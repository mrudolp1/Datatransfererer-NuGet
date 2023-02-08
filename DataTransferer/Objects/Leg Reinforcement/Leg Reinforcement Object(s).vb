Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
'Imports Microsoft.Office.Interop

Partial Public Class LegReinforcement
    Inherits EDSExcelObject


#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String = "LegReinforcements"
    Public Overrides ReadOnly Property EDSTableName As String = "tnx.memb_leg_reinforcement"
    Public Overrides ReadOnly Property templatePath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Leg Reinforcement Tool.xlsm")
    Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
        'Add additional sub table references here. Table names should be consistent with EDS table names. 
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("Leg Reinforcements", "A1:C2", "Details (SAPI)"),
                                            New EXCELDTParameter("Leg Reinforcement Details", "A1:AX201", "Sub Tables (SAPI)")}

            'Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("Leg Reinforcements", "A1:K2", "Details (SAPI)"),
            '                                New EXCELDTParameter("Leg Reinforcement Details", "C2:G18", "Sub Tables (SAPI)"),
            '                                New EXCELDTParameter("Leg Reinforcement Results", "I2:S33", "Results (SAPI)")}

            'note: Excel table names are consistent with EDS table names to limit work required within constructors

        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String

        If _Insert = "" Then
            _Insert = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_INSERT
        End If
        SQLInsert = _Insert

        'Top Level
        SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT FIELDS]", Me.SQLInsertFields)

        'Details
        If Me.LegReinforcementDetails.Count > 0 Then
            SQLInsert = SQLInsert.Replace("--BEGIN --[LEG REINFORCEMENT DETAIL INSERT BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAIL INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[LEG REINFORCEMENT DETAIL INSERT END]", "END --[LEG REINFORCEMENT DETAIL INSERT END]")
            For Each row As LegReinforcementDetail In LegReinforcementDetails
                SQLInsert = SQLInsert.Replace("--[LEG REINFORCEMENT DETAIL INSERT]", row.SQLInsert)
            Next
        End If

        'note: additional insert commands are imbedded within objects sharing similar relationships (e.g. plate details insert located within Connections Object)


        ''Results
        'If Me.Results.Count > 0 Then
        '    SQLInsert = SQLInsert.Replace("--BEGIN --[RESULTS INSERT BEGIN]", "BEGIN --[RESULTS INSERT BEGIN]")
        '    SQLInsert = SQLInsert.Replace("--END --[RESULTS INSERT END]", "END --[RESULTS INSERT END]")
        '    SQLInsert = SQLInsert.Replace("--[RESULTS INSERT]", Me.Results.EDSResultQuery)
        'End If

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String
        'This section not only needs to call update commands but also needs to call insert and delete commands since subtables may involve adding or deleting records

        If _Update = "" Then
            _Update = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_UPDATE
        End If
        SQLUpdate = _Update

        'Top Level
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)

        'Details
        If Me.LegReinforcementDetails.Count > 0 Then
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[LEG REINFORCEMENT DETAIL UPDATE BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAIL UPDATE BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[LEG REINFORCEMENT DETAIL UPDATE END]", "END --[LEG REINFORCEMENT DETAIL UPDATE END]")
            For Each row As LegReinforcementDetail In LegReinforcementDetails
                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                    'following fields include default values and therefore are removed from check below: end_connection_type, applied_load_type, slenderness_ratio_type, print_bolt_on_connections, reinforcement_type
                    If IsSomething(row.leg_load_time_mod_option) Or IsSomething(row.leg_crushing) Or IsSomethingString(row.intermeditate_connection_type) Or IsSomething(row.intermeditate_connection_spacing) Or IsSomething(row.ki_override) Or IsSomething(row.leg_diameter) Or IsSomething(row.leg_thickness) Or IsSomething(row.leg_grade) Or IsSomething(row.leg_unbraced_length) Or IsSomething(row.rein_diameter) Or IsSomething(row.rein_thickness) Or IsSomething(row.rein_grade) Or IsSomething(row.leg_length) Or IsSomething(row.rein_length) Or IsSomething(row.set_top_to_bottom) Or IsSomething(row.flange_bolt_quantity_bot) Or IsSomething(row.flange_bolt_circle_bot) Or IsSomething(row.flange_bolt_orientation_bot) Or IsSomething(row.flange_bolt_quantity_top) Or IsSomething(row.flange_bolt_circle_top) Or IsSomething(row.flange_bolt_orientation_top) Or IsSomethingString(row.threaded_rod_size_bot) Or IsSomethingString(row.threaded_rod_mat_bot) Or IsSomething(row.threaded_rod_quantity_bot) Or IsSomething(row.threaded_rod_unbraced_length_bot) Or IsSomethingString(row.threaded_rod_size_top) Or IsSomethingString(row.threaded_rod_mat_top) Or IsSomething(row.threaded_rod_quantity_top) Or IsSomething(row.threaded_rod_unbraced_length_top) Or IsSomething(row.stiffener_height_bot) Or IsSomething(row.stiffener_length_bot) Or IsSomething(row.stiffener_fillet_bot) Or IsSomething(row.stiffener_exx_bot) Or IsSomething(row.flange_thickness_bot) Or IsSomething(row.stiffener_height_top) Or IsSomething(row.stiffener_length_top) Or IsSomething(row.stiffener_fillet_top) Or IsSomething(row.stiffener_exx_top) Or IsSomething(row.flange_thickness_top) Or IsSomethingString(row.structure_ind) Or IsSomethingString(row.leg_reinforcement_name) Or IsSomething(row.top_elev) Or IsSomething(row.bot_elev) Then
                        SQLUpdate = SQLUpdate.Replace("--[LEG REINFORCEMENT DETAIL INSERT]", row.SQLUpdate)
                    Else
                        SQLUpdate = SQLUpdate.Replace("--[LEG REINFORCEMENT DETAIL INSERT]", row.SQLDelete)
                    End If
                Else
                    SQLUpdate = SQLUpdate.Replace("--[LEG REINFORCEMENT DETAIL INSERT]", row.SQLInsert)
                End If
            Next
        End If

        'note: additional update commands are imbedded within objects sharing similar relationships (e.g. plate details update located within Connections Object)


        ''Results
        'If Me.Results.Count > 0 Then
        '    SQLUpdate = SQLUpdate.Replace("--BEGIN --[RESULTS UPDATE BEGIN]", "BEGIN --[RESULTS UPDATE BEGIN]")
        '    SQLUpdate = SQLUpdate.Replace("--END --[RESULTS UPDATE END]", "END --[RESULTS UPDATE END]")
        '    SQLUpdate = SQLUpdate.Replace("--[RESULTS INSERT]", Me.Results.EDSResultQuery)
        'End If

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'Top Level
        If _Delete = "" Then
            _Delete = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_DELETE
        End If
        SQLDelete = _Delete
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)

        'Details
        If Me.LegReinforcementDetails.Count > 0 Then
            SQLDelete = SQLDelete.Replace("--BEGIN --[LEG REINFORCEMENT DETAIL DELETE BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAIL DELETE BEGIN]")
            SQLDelete = SQLDelete.Replace("--END --[LEG REINFORCEMENT DETAIL DELETE END]", "END --[LEG REINFORCEMENT DETAIL DELETE END]")
            For Each row As LegReinforcementDetail In LegReinforcementDetails
                SQLDelete = SQLDelete.Replace("--[LEG REINFORCEMENT DETAIL INSERT]", row.SQLDelete)
            Next
        End If

        'note: additional delete commands are imbedded within objects sharing similar relationships (e.g. plate details delete located within Connections Object)

        Return SQLDelete

    End Function

#End Region

#Region "Define"

    'Private _ID As Integer? 'Defined in EDSObject
    'Private _tool_version As String 'Defined in EDSExcelObject
    'Private _bus_unit As Integer? 'Defined in EDSObject
    'Private _structure_id As String 'Defined in EDSObject
    'Private _modified_person_id As Integer? 'Defined in EDSExcelObject
    'Private _process_stage As String 'Defined in EDSExcelObject
    Private _Structural_105 As Boolean?

    Public Property LegReinforcementDetails As New List(Of LegReinforcementDetail)

    '<Category("Leg Reinforcements"), Description(""), DisplayName("Id")>
    'Public Property ID() As Integer?
    '    Get
    '        Return Me._ID
    '    End Get
    '    Set
    '        Me._ID = Value
    '    End Set
    'End Property
    '<Category("Leg Reinforcements"), Description(""), DisplayName("Tool Version")>
    'Public Property tool_version() As String
    '    Get
    '        Return Me._tool_version
    '    End Get
    '    Set
    '        Me._tool_version = Value
    '    End Set
    'End Property
    '<Category("Leg Reinforcements"), Description(""), DisplayName("Bus Unit")>
    'Public Property bus_unit() As Integer?
    '    Get
    '        Return Me._bus_unit
    '    End Get
    '    Set
    '        Me._bus_unit = Value
    '    End Set
    'End Property
    '<Category("Leg Reinforcements"), Description(""), DisplayName("Structure Id")>
    'Public Property structure_id() As String
    '    Get
    '        Return Me._structure_id
    '    End Get
    '    Set
    '        Me._structure_id = Value
    '    End Set
    'End Property
    '<Category("Leg Reinforcements"), Description(""), DisplayName("Modified Person Id")>
    'Public Property modified_person_id() As Integer?
    '    Get
    '        Return Me._modified_person_id
    '    End Get
    '    Set
    '        Me._modified_person_id = Value
    '    End Set
    'End Property
    '<Category("Leg Reinforcements"), Description(""), DisplayName("Process Stage")>
    'Public Property process_stage() As String
    '    Get
    '        Return Me._process_stage
    '    End Get
    '    Set
    '        Me._process_stage = Value
    '    End Set
    'End Property
    <Category("Leg Reinforcements"), Description(""), DisplayName("Structural 105")>
    Public Property Structural_105() As Boolean?
        Get
            Return Me._Structural_105
        End Get
        Set
            Me._Structural_105 = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal dr As DataRow, ByRef strDS As DataSet, Optional ByRef Parent As EDSObject = Nothing) 'Added strDS in order to pull EDS data from subtables
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        'Following used to create dataset, regardless if source was EDS or Excel. Boolean used to identify source. EDS = True
        BuildFromDataset(dr, strDS, True, Me)

    End Sub 'Generate Leg Reinforcement from EDS

    Public Sub New(ExcelFilePath As String, Optional ByRef Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet

        For Each item As EXCELDTParameter In excelDTParams
            'Get additional tables from excel file 
            Try
                excelDS.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
            Catch ex As Exception
                Debug.Print(String.Format("Failed to create datatable for: {0}, {1}, {2}", IO.Path.GetFileName(ExcelFilePath), item.xlsSheet, item.xlsRange))
            End Try
        Next

        If excelDS.Tables.Contains("Leg Reinforcements") Then
            Dim dr = excelDS.Tables("Leg Reinforcements").Rows(0)

            'Following used to create dataset, regardless if source was EDS or Excel. Boolean used to identify source. Excel = False
            BuildFromDataset(dr, excelDS, False, Me)

        End If

    End Sub 'Generate Leg Reinforcement from Excel

    Private Sub BuildFromDataset(ByVal dr As DataRow, ByRef ds As DataSet, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing)
        'Dataset is pulled in from either EDS or Excel. True = EDS, False = Excel
        'If Parent IsNot Nothing Then Me.Absorb(Parent) 'Do not double absorb!!!

        'Not sure this is necessary, could just read the values from the structure code criteria when creating the Excel sheet (Added to Save to Excel Section)
        'Me.tia_current = Me.ParentStructure?.structureCodeCriteria?.tia_current
        'Me.rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
        'Me.seismic_design_category = Me.ParentStructure?.structureCodeCriteria?.seismic_design_category

        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.tool_version = DBtoStr(dr.Item("tool_version"))
        Me.bus_unit = If(EDStruefalse, DBtoStr(dr.Item("bus_unit")), Me.bus_unit) 'Not provided in Excel
        Me.work_order_seq_num = If(EDStruefalse, Me.work_order_seq_num, Me.work_order_seq_num) 'Not provided in Excel
        Me.structure_id = If(EDStruefalse, DBtoStr(dr.Item("structure_id")), Me.structure_id) 'Not provided in Excel
        Me.modified_person_id = If(EDStruefalse, DBtoNullableInt(dr.Item("modified_person_id")), Me.modified_person_id) 'Not provided in Excel
        Me.process_stage = If(EDStruefalse, DBtoStr(dr.Item("process_stage")), Me.process_stage) 'Not provided in Excel
        Me.Structural_105 = DBtoNullableBool(dr.Item("Structural_105"))

        Dim lrDetails As New LegReinforcementDetail 'Leg Reinforcement Details
        'Dim plPlateResult As New PlateResults 'Plate Results

        For Each lrrow As DataRow In ds.Tables(lrDetails.EDSObjectName).Rows
            'create a new connection based on the datarow from above
            lrDetails = New LegReinforcementDetail(lrrow, EDStruefalse, Me)
            'Check if the parent id, in the case leg reinforcement id is equal to the original object id (Me)                    
            If If(EDStruefalse, lrDetails.leg_reinforcement_id = Me.ID, True) Then 'If coming from Excel, all leg reinforcement details provided will be associated to leg reinforcement. 
                'If it is equal then add the newly created connection to the list of connections 
                LegReinforcementDetails.Add(lrDetails)
            End If
            'If IsSomething(ds.Tables("Plate Results")) Then
            '                    For Each prrow As DataRow In ds.Tables("Plate Results").Rows
            '                        plPlateResult = New PlateResults(prrow, EDStruefalse, Me)
            '                        If If(EDStruefalse, False, plPlateResult.local_plate_id = plPlateDetail.local_id) Then
            '                            plPlateDetail.PlateResults.Add(plPlateResult)
            '                        End If
            '                    Next
            '                End If
        Next

    End Sub

#End Region

#Region "Save to Excel"

    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        '''''Customize for each excel tool'''''
        Dim LegReinRow As Integer

        'Site Code Criteria
        Dim tia_current, site_name, structure_type As String
        Dim rev_h_section_15_5 As Boolean?

        With wb
            'Site Code Criteria
            'Site Name
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.site_name) Then
                site_name = Me.ParentStructure?.structureCodeCriteria?.site_name
                .Worksheets("IMPORT").Range("SiteName_Import").Value = CType(site_name, String)
            End If
            'Order Number - currently referencing work_order_seq_num below
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.order_number) Then
            '    site_name = Me.ParentStructure?.structureCodeCriteria?.order_number
            '    .Worksheets("IMPORT").Range("Order_Import").Value = CType(order_number, String)
            'End If
            'Tower Type - Defaulting to Self-Support if can't determine
            'Tool is set up where importing tnx file determines tower type. Pulling this from the site code criteria might not be necessary since importing geometry is required.
            ''If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.structure_type) Then
            'If Me.ParentStructure?.structureCodeCriteria?.structure_type = "SELF SUPPORT" Then
            '    structure_type = "Self Support"
            'ElseIf Me.ParentStructure?.structureCodeCriteria?.structure_type = "GUYED TOWER" Then ' ****Note sure if this is correct, need to validate****
            '    structure_type = "Guyed Tower"
            'Else
            '    structure_type = "Self Support"
            'End If
            '.Worksheets("IMPORT").Range("TowerTypeImport").Value = CType(structure_type, String)
            ''End If
            'TIA Revision- Defaulting to Rev. H if not available. Currently pulled in from TNX file
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.tia_current) Then
            '    If Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-F" Then
            '        tia_current = "F"
            '    ElseIf Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-G" Then
            '        tia_current = "G"
            '    ElseIf Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-H" Then
            '        tia_current = "H"
            '    Else
            '        tia_current = "H"
            '    End If
            '    .Worksheets("IMPORT").Range("TIA_Import").Value = CType(tia_current, String)
            'End If
            'H Section 15.5 - Not sure if this is a reliable source. Currently just pulling from last SA (Structural_105)
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5) Then
            '    rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
            '    .Worksheets("IMPORT").Range("U1").Value = CType(rev_h_section_15_5, Boolean)
            'End If
            'Load Z Normalization
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.load_z_norm) Then
            '    rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.load_z_norm
            '    .Worksheets("Engine").Range("G10").Value = CType(load_z_norm, Boolean)
            'End If

            'Loading
            'If structure_type = "Self-Suppot" Then
            '    .Worksheets("Input").Range("D13").Value = CType(uplift, Double)
            '    .Worksheets("Input").Range("D14").Value = CType(compression, Double)
            '    .Worksheets("Input").Range("D15").Value = CType(uplift_shear, Double)
            '    .Worksheets("Input").Range("D16").Value = CType(compression_shear, Double)
            'Else
            '    .Worksheets("Input").Range("D13").Value = CType(moment, Double)
            '    .Worksheets("Input").Range("D14").Value = CType(axial, Double)
            '    .Worksheets("Input").Range("D15").Value = CType(shear, Double)
            'End If

            .Worksheets("Details (SAPI)").Range("A3").Value = CType(True, Boolean) 'Flags if sheet was last touched by EDS. If true, worksheet change event upon opening tool. 

            .Worksheets("Details (SAPI)").Range("ID").Value = CType(Me.ID, Integer)
            ''If Not IsNothing(Me.ID) Then
            ''    .Worksheets("Sub Tables (SAPI)").Range("ID").Value = CType(Me.ID, Integer)
            ''Else
            ''    .Worksheets("Sub Tables (SAPI)").Range("ID").ClearContents
            ''End If
            'If Not IsNothing(Me.tool_version) Then
            '    .Worksheets("Reference").Range("C23").Value = CType(Me.tool_version, String)
            'End If
            If Not IsNothing(Me.bus_unit) Then
                .Worksheets("IMPORT").Range("BU_Import").Value = CType(Me.bus_unit, Integer)
            Else
                .Worksheets("IMPORT").Range("BU_Import").ClearContents
            End If
            If Not IsNothing(Me.work_order_seq_num) Then
                .Worksheets("IMPORT").Range("Order_Import").Value = CType(Me.work_order_seq_num, Integer)
            Else
                .Worksheets("IMPORT").Range("Order_Import").ClearContents
            End If
            'If Not IsNothing(Me.structure_id) Then
            '    .Worksheets("").Range("").Value = CType(Me.structure_id, String)
            'End If
            'If Not IsNothing(Me.modified_person_id) Then
            '    .Worksheets("").Range("").Value = CType(Me.modified_person_id, Integer)
            'Else
            '    .Worksheets("").Range("").ClearContents
            'End If
            'If Not IsNothing(Me.process_stage) Then
            '    .Worksheets("").Range("").Value = CType(Me.process_stage, String)
            'End If
            If Not IsNothing(Me.Structural_105) Then
                .Worksheets("Summary").Range("U1").Value = CType(Me.Structural_105, Boolean)
            End If

            If Me.LegReinforcementDetails.Count > 0 Then
                For Each row As LegReinforcementDetail In LegReinforcementDetails
                    LegReinRow = CType(row.local_id, Integer) + 3 'determines row to enter data in within tool
                    'Select which rows have reinforcement on the IMPORT tab
                    If Not IsNothing(LegReinRow) Then
                        .Worksheets("IMPORT").Range("H" & LegReinRow + 13).Value = "Yes"
                    Else
                        .Worksheets("IMPORT").Range("H" & LegReinRow + 13).ClearContents
                    End If
            If Not IsNothing(row.ID) Then
                        .Worksheets("Sub Tables (SAPI)").Range("B" & LegReinRow - 2).Value = CType(row.ID, Integer)
                    Else
                        .Worksheets("Sub Tables (SAPI)").Range("B" & LegReinRow - 2).ClearContents
                    End If
                    'If Not IsNothing(row.leg_reinforcement_id) Then
                    '    .Worksheets("Databse").Range("").Value = CType(row.leg_reinforcement_id, Integer)
                    'Else
                    '    .Worksheets("Databse").Range("").ClearContents
                    'End If
                    If Not IsNothing(row.leg_load_time_mod_option) Then
                        If row.leg_load_time_mod_option = True Then
                            .Worksheets("Database").Range("AB" & LegReinRow).Value = "Yes"
                        Else
                            .Worksheets("Database").Range("AB" & LegReinRow).Value = "No"
                        End If
                    End If
                    If Not IsNothing(row.end_connection_type) Then
                        .Worksheets("Database").Range("AC" & LegReinRow).Value = CType(row.end_connection_type, String)
                    End If
                    If Not IsNothing(row.leg_crushing) Then
                        If row.leg_crushing = True Then
                            .Worksheets("Database").Range("AD" & LegReinRow).Value = "Yes"
                        Else
                            .Worksheets("Database").Range("AD" & LegReinRow).Value = "No"
                        End If
                    End If
                    If Not IsNothing(row.applied_load_type) Then
                        .Worksheets("Database").Range("AE" & LegReinRow).Value = CType(row.applied_load_type, String)
                    End If
                    If Not IsNothing(row.slenderness_ratio_type) Then
                        .Worksheets("Database").Range("AF" & LegReinRow).Value = CType(row.slenderness_ratio_type, String)
                    End If
                    If Not IsNothing(row.intermeditate_connection_type) Then
                        .Worksheets("Database").Range("AG" & LegReinRow).Value = CType(row.intermeditate_connection_type, String)
                    End If
                    If Not IsNothing(row.intermeditate_connection_spacing) Then
                        .Worksheets("Database").Range("AH" & LegReinRow).Value = CType(row.intermeditate_connection_spacing, Double)
                    Else
                        .Worksheets("Database").Range("AH" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.ki_override) Then
                        .Worksheets("Database").Range("AI" & LegReinRow).Value = CType(row.ki_override, Double)
                    Else
                        .Worksheets("Database").Range("AI" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.leg_diameter) Then
                        .Worksheets("Database").Range("AJ" & LegReinRow).Value = CType(row.leg_diameter, Double)
                    Else
                        .Worksheets("Database").Range("AJ" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.leg_thickness) Then
                        .Worksheets("Database").Range("AK" & LegReinRow).Value = CType(row.leg_thickness, Double)
                    Else
                        .Worksheets("Database").Range("AK" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.leg_grade) Then
                        .Worksheets("Database").Range("AL" & LegReinRow).Value = CType(row.leg_grade, Double)
                    Else
                        .Worksheets("Database").Range("AL" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.leg_unbraced_length) Then
                        .Worksheets("Database").Range("AM" & LegReinRow).Value = CType(row.leg_unbraced_length, Double)
                    Else
                        .Worksheets("Database").Range("AM" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.rein_diameter) Then
                        .Worksheets("Database").Range("AN" & LegReinRow).Value = CType(row.rein_diameter, Double)
                    Else
                        .Worksheets("Database").Range("AN" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.rein_thickness) Then
                        .Worksheets("Database").Range("AO" & LegReinRow).Value = CType(row.rein_thickness, Double)
                    Else
                        .Worksheets("Database").Range("AO" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.rein_grade) Then
                        .Worksheets("Database").Range("AP" & LegReinRow).Value = CType(row.rein_grade, Double)
                    Else
                        .Worksheets("Database").Range("AP" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.print_bolt_on_connections) Then
                        .Worksheets("Database").Range("AS" & LegReinRow).Value = CType(row.print_bolt_on_connections, Boolean)
                    End If
                    If Not IsNothing(row.leg_length) Then
                        .Worksheets("Database").Range("AT" & LegReinRow).Value = CType(row.leg_length, Double)
                    Else
                        .Worksheets("Database").Range("AT" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.rein_length) Then
                        .Worksheets("Database").Range("AU" & LegReinRow).Value = CType(row.rein_length, Double)
                    Else
                        .Worksheets("Database").Range("AU" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.set_top_to_bottom) Then
                        .Worksheets("Database").Range("AV" & LegReinRow).Value = CType(row.set_top_to_bottom, Boolean)
                    End If
                    If Not IsNothing(row.flange_bolt_quantity_bot) Then
                        .Worksheets("Database").Range("AW" & LegReinRow).Value = CType(row.flange_bolt_quantity_bot, Integer)
                    Else
                        .Worksheets("Database").Range("AW" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.flange_bolt_circle_bot) Then
                        .Worksheets("Database").Range("AX" & LegReinRow).Value = CType(row.flange_bolt_circle_bot, Double)
                    Else
                        .Worksheets("Database").Range("AX" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.flange_bolt_orientation_bot) Then
                        .Worksheets("Database").Range("AY" & LegReinRow).Value = CType(row.flange_bolt_orientation_bot, Integer)
                    Else
                        .Worksheets("Database").Range("AY" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.flange_bolt_quantity_top) Then
                        .Worksheets("Database").Range("AZ" & LegReinRow).Value = CType(row.flange_bolt_quantity_top, Integer)
                    Else
                        .Worksheets("Database").Range("AZ" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.flange_bolt_circle_top) Then
                        .Worksheets("Database").Range("BA" & LegReinRow).Value = CType(row.flange_bolt_circle_top, Double)
                    Else
                        .Worksheets("Database").Range("BA" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.flange_bolt_orientation_top) Then
                        .Worksheets("Database").Range("BB" & LegReinRow).Value = CType(row.flange_bolt_orientation_top, Integer)
                    Else
                        .Worksheets("Database").Range("BB" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.threaded_rod_size_bot) Then
                        .Worksheets("Database").Range("BC" & LegReinRow).Value = CType(row.threaded_rod_size_bot, String)
                    End If
                    If Not IsNothing(row.threaded_rod_mat_bot) Then
                        .Worksheets("Database").Range("BD" & LegReinRow).Value = CType(row.threaded_rod_mat_bot, String)
                    End If
                    If Not IsNothing(row.threaded_rod_quantity_bot) Then
                        .Worksheets("Database").Range("BE" & LegReinRow).Value = CType(row.threaded_rod_quantity_bot, Integer)
                    Else
                        .Worksheets("Database").Range("BE" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.threaded_rod_unbraced_length_bot) Then
                        .Worksheets("Database").Range("BF" & LegReinRow).Value = CType(row.threaded_rod_unbraced_length_bot, Double)
                    Else
                        .Worksheets("Database").Range("BF" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.threaded_rod_size_top) Then
                        .Worksheets("Database").Range("BG" & LegReinRow).Value = CType(row.threaded_rod_size_top, String)
                    End If
                    If Not IsNothing(row.threaded_rod_mat_top) Then
                        .Worksheets("Database").Range("BH" & LegReinRow).Value = CType(row.threaded_rod_mat_top, String)
                    End If
                    If Not IsNothing(row.threaded_rod_quantity_top) Then
                        .Worksheets("Database").Range("BI" & LegReinRow).Value = CType(row.threaded_rod_quantity_top, Integer)
                    Else
                        .Worksheets("Database").Range("BI" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.threaded_rod_unbraced_length_top) Then
                        .Worksheets("Database").Range("BJ" & LegReinRow).Value = CType(row.threaded_rod_unbraced_length_top, Double)
                    Else
                        .Worksheets("Database").Range("BJ" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.stiffener_height_bot) Then
                        .Worksheets("Database").Range("BK" & LegReinRow).Value = CType(row.stiffener_height_bot, Double)
                    Else
                        .Worksheets("Database").Range("BK" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.stiffener_length_bot) Then
                        .Worksheets("Database").Range("BL" & LegReinRow).Value = CType(row.stiffener_length_bot, Double)
                    Else
                        .Worksheets("Database").Range("BL" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.stiffener_fillet_bot) Then
                        .Worksheets("Database").Range("BM" & LegReinRow).Value = CType(row.stiffener_fillet_bot, Integer)
                    Else
                        .Worksheets("Database").Range("BM" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.stiffener_exx_bot) Then
                        .Worksheets("Database").Range("BN" & LegReinRow).Value = CType(row.stiffener_exx_bot, Double)
                    Else
                        .Worksheets("Database").Range("BN" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.flange_thickness_bot) Then
                        .Worksheets("Database").Range("BO" & LegReinRow).Value = CType(row.flange_thickness_bot, Double)
                    Else
                        .Worksheets("Database").Range("BO" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.stiffener_height_top) Then
                        .Worksheets("Database").Range("BP" & LegReinRow).Value = CType(row.stiffener_height_top, Double)
                    Else
                        .Worksheets("Database").Range("BP" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.stiffener_length_top) Then
                        .Worksheets("Database").Range("BQ" & LegReinRow).Value = CType(row.stiffener_length_top, Double)
                    Else
                        .Worksheets("Database").Range("BQ" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.stiffener_fillet_top) Then
                        .Worksheets("Database").Range("BR" & LegReinRow).Value = CType(row.stiffener_fillet_top, Integer)
                    Else
                        .Worksheets("Database").Range("BR" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.stiffener_exx_top) Then
                        .Worksheets("Database").Range("BS" & LegReinRow).Value = CType(row.stiffener_exx_top, Double)
                    Else
                        .Worksheets("Database").Range("BS" & LegReinRow).ClearContents
                    End If
                    If Not IsNothing(row.flange_thickness_top) Then
                        .Worksheets("Database").Range("BT" & LegReinRow).Value = CType(row.flange_thickness_top, Double)
                    Else
                        .Worksheets("Database").Range("BT" & LegReinRow).ClearContents
                    End If
                    'need to see if this is necessary
                    'If Not IsNothing(row.structure_ind) Then
                    '    .Worksheets("").Range("").Value = CType(row.structure_ind, String)
                    'End If
                    If Not IsNothing(row.reinforcement_type) Then
                        .Worksheets("Database").Range("Y" & LegReinRow).Value = CType(row.reinforcement_type, String)
                    End If
                    'Don't need to store back in tool since will generate based on all previous inputs. 
                    'If Not IsNothing(row.leg_reinforcement_name) Then
                    '    .Worksheets("Database").Range("BV" & LegReinRow).Value = CType(row.leg_reinforcement_name, String)
                    'End If
                    'If Not IsNothing(Me.local_id) Then
                    '    .Worksheets("").Range("").Value = CType(Me.local_id, Integer)
                    'Else
                    '    .Worksheets("").Range("").ClearContents
                    'End If
                    'elevations are not required since they are pulled in with tnx import
                    'If Not IsNothing(row.top_elev) Then
                    '    .Worksheets("Database").Range("I" & LegReinRow).Value = CType(row.top_elev, Double)
                    'Else
                    '    .Worksheets("Database").Range("I" & LegReinRow).ClearContents
                    'End If
                    'If Not IsNothing(row.bot_elev) Then
                    '    .Worksheets("Database").Range("J" & LegReinRow).Value = CType(row.bot_elev, Double)
                    'Else
                    '    .Worksheets("Database").Range("J" & LegReinRow).ClearContents
                    'End If

                Next

            End If



            'If Me.Connections.Count > 0 Then
            '    'identify first row to copy data into Excel Sheet
            '    'Connection
            '    Dim PlateRow As Integer = 3 'SAPI Tab
            '    Dim PlateRow2 As Integer = 46 'MP Connection Summary Tab
            '    'Dim PlateRow3 As Integer = 19 'Main Tab (currently not used)
            '    'Plate Details
            '    Dim PlateDRow As Integer = 3 'SAPI Tab
            '    Dim i As Integer 'Row adjustment for top vs bottom plate
            '    'Materials
            '    Dim MatRow As Integer = 40 'Materials Tab; SAPI Tab - 2
            '    Dim tempMaterials As New List(Of CCIplateMaterial)
            '    Dim tempMaterial As New CCIplateMaterial 'Temp material object to determine if already added to Excel
            '    Dim matflag As Boolean = False 'determines whether or not to add to Excel based on temp list
            '    'Excel Database Reference
            '    Dim mycol As Integer = 6 'Bolt Group, Bolt Details, Stiffener Details
            '    Dim bump As Integer = 0 'Bolt Group
            '    Dim bump2 As Integer = 0 'Bolt Detail
            '    'Stiffener Group
            '    Dim StiffGRow As Integer = 3 'SAPI Tab
            '    'Dim StiffGRow2 As Integer = 10 'MP Connection Summary TAB
            '    'Stiffener Details
            '    Dim StiffDRow As Integer = 3 'SAPI Tab
            '    Dim StiffDRow2 As Integer = 85 ' MP Connection Summary Tab
            '    'Bridge Stiffener Details
            '    Dim BridgeDRow As Integer = 3 'SAPI Tab
            '    'Dim BridgeDRow2 As Integer = 167 'MP Connection Summary Tab

            '    For Each row As Connection In Connections

            '        'Excel Database Reference (resets for each plate connection)
            '        Dim myrow4 As Integer '= 1027 'Stiffener Details

            '        If Not IsNothing(row.ID) Then
            '            .Worksheets("Sub Tables (SAPI)").Range("D" & PlateRow).Value = CType(row.ID, Integer)
            '        End If
            '        If structure_type = "Self Support" Then
            '            If Not IsNothing(row.connection_elevation) Then
            '                .Worksheets("Custom Connection").Range("elevation").Value = CType(row.connection_elevation, Double)
            '            End If
            '            If Not IsNothing(row.bolt_configuration) Then
            '                .Worksheets("Main").Range("D246").Value = CType(row.bolt_configuration, String)
            '            Else
            '                .Worksheets("Main").Range("D246").ClearContents
            '            End If
            '        ElseIf structure_type = "Monopole" Then
            '            If Not IsNothing(row.connection_elevation) Then
            '                .Worksheets("MP Connection Summary").Range("C" & PlateRow2).Value = CType(row.connection_elevation, Double)
            '            End If
            '            If Not IsNothing(row.bolt_configuration) Then
            '                .Worksheets("MP Connection Summary").Range("K" & PlateRow2).Value = CType(row.bolt_configuration, String)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("K" & PlateRow2).ClearContents
            '            End If
            '            'Need to report flags for proper worksheet change events
            '            If row.bolt_configuration = "Custom" And row.connection_type = "Base" Then
            '                .Worksheets("MP Connection Summary").Range("U" & PlateRow2).Value = CType(1, Integer)
            '            ElseIf row.bolt_configuration = "Custom" And row.connection_type = "Flange" Then
            '                .Worksheets("MP Connection Summary").Range("U" & PlateRow2).Value = CType(1, Integer)
            '                .Worksheets("MP Connection Summary").Range("U" & PlateRow2 + 1).Value = CType(1, Integer)
            '            End If
            '        End If
            '        'If Not IsNothing(row.connection_type) Then 'do not need, tool will autopopulate
            '    '    .Worksheets("").Range("").Value = CType(row.connection_type, String)
            '    'End If

            '    'For i = 1 To 200 '200 possilbe rows for Pole Geometry (Need to figure out how to add. Need to reference pole geometry)
            '    '    'If row.connection_elevation = .Worksheets("Main").Range("B" & PlateRow3).Value Then
            '    '    If Not IsNothing(row.connection_type) Then
            '    '        .Worksheets("Main").Range("B" & PlateRow3).Value = CType(row.connection_type, String)
            '    '    Else
            '    '        .Worksheets("Main").Range("B" & PlateRow3).ClearContents
            '    '    End If
            '    '    'End If
            '    'Next i

            '    'If Not IsNothing(row.bolt_configuration) Then
            '    '    If structure_type = "Self Support" Then
            '    '        .Worksheets("Main").Range("D246").Value = CType(row.bolt_configuration, String)
            '    '    ElseIf structure_type = "Monopole" Then
            '    '        .Worksheets("MP Connection Summary").Range("K" & PlateRow2).Value = CType(row.bolt_configuration, String)
            '    '        'Need to report flags for proper worksheet change events
            '    '        If row.bolt_configuration = "Custom" And row.connection_type = "Base" Then
            '    '            .Worksheets("MP Connection Summary").Range("U" & PlateRow2).Value = CType(1, Integer)
            '    '        ElseIf row.bolt_configuration = "Custom" And row.connection_type = "Flange" Then
            '    '            .Worksheets("MP Connection Summary").Range("U" & PlateRow2).Value = CType(1, Integer)
            '    '            .Worksheets("MP Connection Summary").Range("U" & PlateRow2 + 1).Value = CType(1, Integer)
            '    '        End If
            '    '    End If
            '    'End If

            '    For Each pdrow As PlateDetail In row.PlateDetails



            '        If pdrow.connection_id = row.ID Then

            '            If pdrow.plate_location = "Bottom" Then
            '                i = 1
            '                myrow4 = 2427
            '            Else
            '                i = 0
            '                myrow4 = 1027
            '            End If

            '            If Not IsNothing(pdrow.ID) Then
            '                .Worksheets("Sub Tables (SAPI)").Range("K" & PlateDRow).Value = CType(pdrow.ID, Integer)
            '            End If
            '            'If Not IsNothing(row.plate_location) Then 'do not need, tool will autopopulate
            '            '    .Worksheets("MP Connection Summary").Range("D" & PlateRow2).Value = CType(row.plate_location, String)
            '            'End If
            '            If Not IsNothing(pdrow.plate_type) Then
            '                .Worksheets("MP Connection Summary").Range("E" & PlateRow2 + i).Value = CType(pdrow.plate_type, String)
            '            End If
            '            If Not IsNothing(pdrow.plate_diameter) Then
            '                .Worksheets("MP Connection Summary").Range("F" & PlateRow2 + i).Value = CType(pdrow.plate_diameter, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("F" & PlateRow2 + i).ClearContents
            '            End If
            '            If Not IsNothing(pdrow.plate_thickness) Then
            '                .Worksheets("MP Connection Summary").Range("G" & PlateRow2 + i).Value = CType(pdrow.plate_thickness, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("G" & PlateRow2 + i).ClearContents
            '            End If
            '            'If Not IsNothing(pdrow.plate_material) Then
            '            '    .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).Value = CType(pdrow.plate_material, Integer)
            '            'Else
            '            '    .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).ClearContents
            '            'End If
            '            For Each mrow As CCIplateMaterial In pdrow.CCIplateMaterials
            '                If mrow.default_material = True Then
            '                    If Not IsNothing(mrow.name) Then
            '                        .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).Value = CType(mrow.name, String)
            '                    Else
            '                        .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).ClearContents
            '                    End If
            '                Else 'After adding new materail, save material name in a list to reference for other plates to see if materail was already added. 
            '                    For Each tmrow In tempMaterials
            '                        If mrow.ID = tmrow.ID Then
            '                            matflag = True 'don't add to excel
            '                            Exit For
            '                        End If
            '                    Next
            '                    If matflag = False Then
            '                        tempMaterial = New CCIplateMaterial(mrow.ID)
            '                        tempMaterials.Add(tempMaterial)

            '                        '.Worksheets("Sub Tables (SAPI)").Range("AR").Count
            '                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Columns("AR").Count 'counts total columns in excel
            '                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Cells("AR:AR").Count
            '                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").GetDataRange("AR:AR")

            '                        'If Not IsNothing(mrow.ID) Then
            '                        '    .Worksheets("Sub Tables (SAPI)").Range("AR" & MatRow2).Value = CType(mrow.ID, Integer)
            '                        'End If
            '                        If Not IsNothing(mrow.name) Then
            '                            .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).Value = CType(mrow.name, String)
            '                        Else
            '                            .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).ClearContents
            '                        End If

            '                        SaveMaterial(wb, mrow, MatRow)
            '                        MatRow += 1
            '                    Else
            '                        If Not IsNothing(mrow.name) Then
            '                            .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).Value = CType(mrow.name, String)
            '                        Else
            '                            .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).ClearContents
            '                        End If
            '                    End If
            '                End If
            '                matflag = False 'reset flag
            '            Next

            '            If Not IsNothing(pdrow.stiffener_configuration) Then
            '                If pdrow.stiffener_configuration = 4 Then
            '                    .Worksheets("MP Connection Summary").Range("I" & PlateRow2 + i).Value = "Custom"
            '                    'Need to report flags for proper worksheet change events
            '                    .Worksheets("MP Connection Summary").Range("T" & PlateRow2).Value = CType(1, Integer)
            '                Else
            '                    .Worksheets("MP Connection Summary").Range("I" & PlateRow2 + i).Value = CType(pdrow.stiffener_configuration, Integer)
            '                End If
            '            Else
            '                .Worksheets("MP Connection Summary").Range("I" & PlateRow2 + i).ClearContents
            '            End If
            '            If Not IsNothing(pdrow.stiffener_clear_space) Then
            '                .Worksheets("MP Connection Summary").Range("D" & PlateRow2 + i + 39).Value = CType(pdrow.stiffener_clear_space, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("D" & PlateRow2 + i + 39).ClearContents
            '            End If
            '            If Not IsNothing(pdrow.plate_check) Then
            '                If pdrow.plate_check = True Then
            '                    .Worksheets("MP Connection Summary").Range("J" & PlateRow2 + i).Value = "Yes"
            '                Else
            '                    .Worksheets("MP Connection Summary").Range("J" & PlateRow2 + i).Value = "No"
            '                End If
            '            End If

            '            Dim sgid As Integer = 1 'Stiffener Group names per CCIplate are integers 1-5
            '            Dim FirstRowStiff As Boolean = True
            '            'Stiffener Group
            '            StiffGRow = 3 'SAPI Tab -reset for each plate detail
            '            'Stiffener Details
            '            StiffDRow = 3 'SAPI Tab -reset for each plate detail

            '            For Each sgrow As StiffenerGroup In pdrow.StiffenerGroups
            '                If sgrow.plate_details_id = pdrow.ID Then

            '                    If Not IsNothing(sgrow.ID) Then
            '                        .Worksheets("Sub Tables (SAPI)").Range("BI" & StiffGRow + (PlateDRow - 3) * 5).Value = CType(sgrow.ID, Integer)
            '                    End If
            '                    'If Not IsNothing(Me.stiffener_name) Then
            '                    '    .Worksheets("Database").Range("G").Value = CType(Me.stiffener_name, String)
            '                    'End If

            '                    For Each sdrow As StiffenerDetail In sgrow.StiffenerDetails
            '                        If sdrow.stiffener_id = sgrow.ID Then

            '                            'Save stiffener data to MP Connection Summary when connection is symmetrical
            '                            If FirstRowStiff And pdrow.stiffener_configuration > 0 And pdrow.stiffener_configuration <> 4 Then
            '                                'If Not IsNothing(sdrow.stiffener_location) Then
            '                                '    .Worksheets("Database").Cells(myrow4 + 1, mycol).Value = CType(sdrow.stiffener_location, Double)
            '                                'End If
            '                                If Not IsNothing(sdrow.stiffener_width) Then
            '                                    .Worksheets("MP Connection Summary").Range("E" & StiffDRow2).Value = CType(sdrow.stiffener_width, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.stiffener_height) Then
            '                                    .Worksheets("MP Connection Summary").Range("F" & StiffDRow2).Value = CType(sdrow.stiffener_height, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.stiffener_thickness) Then
            '                                    .Worksheets("MP Connection Summary").Range("G" & StiffDRow2).Value = CType(sdrow.stiffener_thickness, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.stiffener_h_notch) Then
            '                                    .Worksheets("MP Connection Summary").Range("H" & StiffDRow2).Value = CType(sdrow.stiffener_h_notch, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.stiffener_v_notch) Then
            '                                    .Worksheets("MP Connection Summary").Range("I" & StiffDRow2).Value = CType(sdrow.stiffener_v_notch, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.stiffener_grade) Then
            '                                    .Worksheets("MP Connection Summary").Range("J" & StiffDRow2).Value = CType(sdrow.stiffener_grade, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.weld_type) Then
            '                                    .Worksheets("MP Connection Summary").Range("K" & StiffDRow2).Value = CType(sdrow.weld_type, String)
            '                                End If
            '                                If Not IsNothing(sdrow.groove_depth) Then
            '                                    .Worksheets("MP Connection Summary").Range("L" & StiffDRow2).Value = CType(sdrow.groove_depth, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.groove_angle) Then
            '                                    .Worksheets("MP Connection Summary").Range("M" & StiffDRow2).Value = CType(sdrow.groove_angle, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.h_fillet_weld) Then
            '                                    .Worksheets("MP Connection Summary").Range("N" & StiffDRow2).Value = CType(sdrow.h_fillet_weld, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.v_fillet_weld) Then
            '                                    .Worksheets("MP Connection Summary").Range("O" & StiffDRow2).Value = CType(sdrow.v_fillet_weld, Double)
            '                                End If
            '                                If Not IsNothing(sdrow.weld_strength) Then
            '                                    .Worksheets("MP Connection Summary").Range("P" & StiffDRow2).Value = CType(sdrow.weld_strength, Double)
            '                                End If

            '                                FirstRowStiff = False
            '                            End If


            '                            If Not IsNothing(sdrow.ID) Then
            '                                .Worksheets("Sub Tables (SAPI)").Range("BN" & StiffDRow + (PlateDRow - 3) * 100).Value = CType(sdrow.ID, Integer)
            '                            Else
            '                                .Worksheets("Sub Tables (SAPI)").Range("BN" & StiffDRow + (PlateDRow - 3) * 100).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.stiffener_id) Then
            '                                .Worksheets("Sub Tables (SAPI)").Range("BO" & StiffDRow + (PlateDRow - 3) * 100).Value = CType(sdrow.stiffener_id, Integer)
            '                                .Worksheets("Database").Cells(myrow4, mycol).Value = CType(sgid, Double)
            '                            Else
            '                                .Worksheets("Sub Tables (SAPI)").Range("BO" & StiffDRow + (PlateDRow - 3) * 100).ClearContents
            '                                .Worksheets("Database").Cells(myrow4, mycol).ClearContents
            '                            End If



            '                            If Not IsNothing(sdrow.stiffener_location) Then
            '                                .Worksheets("Database").Cells(myrow4 + 1, mycol).Value = CType(sdrow.stiffener_location, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 1, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.stiffener_width) Then
            '                                .Worksheets("Database").Cells(myrow4 + 2, mycol).Value = CType(sdrow.stiffener_width, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 2, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.stiffener_height) Then
            '                                .Worksheets("Database").Cells(myrow4 + 3, mycol).Value = CType(sdrow.stiffener_height, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 3, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.stiffener_thickness) Then
            '                                .Worksheets("Database").Cells(myrow4 + 4, mycol).Value = CType(sdrow.stiffener_thickness, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 4, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.stiffener_h_notch) Then
            '                                .Worksheets("Database").Cells(myrow4 + 5, mycol).Value = CType(sdrow.stiffener_h_notch, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 5, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.stiffener_v_notch) Then
            '                                .Worksheets("Database").Cells(myrow4 + 6, mycol).Value = CType(sdrow.stiffener_v_notch, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 6, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.stiffener_grade) Then
            '                                .Worksheets("Database").Cells(myrow4 + 7, mycol).Value = CType(sdrow.stiffener_grade, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 7, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.weld_type) Then
            '                                .Worksheets("Database").Cells(myrow4 + 8, mycol).Value = CType(sdrow.weld_type, String)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 8, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.groove_depth) Then
            '                                .Worksheets("Database").Cells(myrow4 + 9, mycol).Value = CType(sdrow.groove_depth, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 9, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.groove_angle) Then
            '                                .Worksheets("Database").Cells(myrow4 + 10, mycol).Value = CType(sdrow.groove_angle, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 10, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.h_fillet_weld) Then
            '                                .Worksheets("Database").Cells(myrow4 + 11, mycol).Value = CType(sdrow.h_fillet_weld, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 11, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.v_fillet_weld) Then
            '                                .Worksheets("Database").Cells(myrow4 + 12, mycol).Value = CType(sdrow.v_fillet_weld, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 12, mycol).ClearContents
            '                            End If
            '                            If Not IsNothing(sdrow.weld_strength) Then
            '                                .Worksheets("Database").Cells(myrow4 + 13, mycol).Value = CType(sdrow.weld_strength, Double)
            '                            Else
            '                                .Worksheets("Database").Cells(myrow4 + 13, mycol).ClearContents
            '                            End If

            '                            myrow4 += 14
            '                            StiffDRow += 1
            '                            'FirstRow = False 'Turns off saving bolt information to MP Connection Summary Tab if not first row

            '                        End If
            '                    Next
            '                    sgid += 1
            '                    StiffGRow += 1
            '                End If
            '            Next


            '            PlateDRow += 1

            '        End If
            '        StiffDRow2 -= 1
            '    Next

            '    'Excel Database Reference (resets for each plate connection)
            '    Dim myrow As Integer = 7 'Bolt Group & Bolt Details
            '    Dim bgid As Integer = 1 'Bolt Group names per CCIplate are integers 1-5
            '    'Bolt Group
            '    Dim myrow3 As Integer = 3827 'Bolt Group BARB Elevation
            '    Dim BoltGRow As Integer = 3 'SAPI Tab
            '    'Bolt Detail
            '    Dim myrow2 As Integer = 27 'Excel Database
            '    Dim BoltDRow As Integer = 3 'SAPI Tab
            '    Dim FirstRow As Boolean = True 'Apllies for only 1 bolt group that is symmetrical. Saves data to MP Connection Summary Tab

            '    For Each bgrow As BoltGroup In row.BoltGroups
            '        If bgrow.connection_id = row.ID Then

            '            If FirstRow And row.bolt_configuration = "Symmetrical" Then
            '                If structure_type = "Self Support" Then
            '                    If Not IsNothing(bgrow.grout_considered) Then
            '                        If bgrow.grout_considered = True Then
            '                            .Worksheets("Main").Range("D251").Value = "Yes"
            '                        Else
            '                            .Worksheets("Main").Range("D251").Value = "No"
            '                        End If
            '                    End If
            '                ElseIf structure_type = "Monopole" And row.connection_type = "Base" Then
            '                    If Not IsNothing(bgrow.grout_considered) Then
            '                        If bgrow.grout_considered = True Then
            '                            .Worksheets("MP Connection Summary").Range("N9").Value = "Yes"
            '                        Else
            '                            .Worksheets("MP Connection Summary").Range("N9").Value = "No"
            '                        End If
            '                    End If
            '                End If
            '            ElseIf FirstRow And row.bolt_configuration = "Custom" Then
            '                If structure_type = "Self Support" Then
            '                    .Worksheets("Main").Range("D251").ClearContents
            '                End If
            '            End If

            '            If Not IsNothing(bgrow.ID) Then
            '                .Worksheets("Sub Tables (SAPI)").Range("W" & BoltGRow + bump).Value = CType(bgrow.ID, Integer)
            '            End If

            '            If Not IsNothing(bgrow.resist_axial) Then
            '                If bgrow.resist_axial = True Then
            '                    .Worksheets("Database").Cells(myrow, mycol).Value = "Yes"
            '                Else
            '                    .Worksheets("Database").Cells(myrow, mycol).Value = "No"
            '                End If
            '            End If
            '            If Not IsNothing(bgrow.resist_shear) Then
            '                If bgrow.resist_shear = True Then
            '                    .Worksheets("Database").Cells(myrow + 1, mycol).Value = "Yes"
            '                Else
            '                    .Worksheets("Database").Cells(myrow + 1, mycol).Value = "No"
            '                End If
            '            End If
            '            If Not IsNothing(bgrow.plate_bending) Then
            '                If bgrow.plate_bending = True Then
            '                    .Worksheets("Database").Cells(myrow + 2, mycol).Value = "Yes"
            '                Else
            '                    .Worksheets("Database").Cells(myrow + 2, mycol).Value = "No"
            '                End If
            '            End If
            '            If Not IsNothing(bgrow.grout_considered) Then
            '                If bgrow.grout_considered = True Then
            '                    .Worksheets("Database").Cells(myrow + 3, mycol).Value = "Yes"
            '                Else
            '                    .Worksheets("Database").Cells(myrow + 3, mycol).Value = "No"
            '                End If
            '            End If
            '            If Not IsNothing(bgrow.apply_barb_elevation) Then
            '                If bgrow.apply_barb_elevation = True Then
            '                    .Worksheets("Database").Cells(myrow3, mycol).Value = "Yes"
            '                Else
            '                    .Worksheets("Database").Cells(myrow3, mycol).Value = "No"
            '                End If
            '            End If
            '            'Bolt Group names in CCIplate are named 1 through 5. 
            '            'If Not IsNothing(bgrow.bolt_name) Then
            '            '    .Worksheets("Database").Cells(myrow + 5, mycol).Value = CType(bgrow.bolt_name, String)
            '            'End If

            '            For Each bdrow As BoltDetail In bgrow.BoltDetails
            '                If bdrow.bolt_group_id = bgrow.ID Then

            '                    'Save bolt data to MP Connection Summary when connection is symmetrical
            '                    If FirstRow And row.bolt_configuration = "Symmetrical" Then
            '                        If structure_type = "Self Support" Then
            '                            If Not IsNothing(bgrow.BoltDetails.Count) Then
            '                                .Worksheets("Main").Range("D248").Value = CType(bgrow.BoltDetails.Count, Integer)
            '                            End If
            '                            If Not IsNothing(bdrow.bolt_diameter) Then
            '                                .Worksheets("Main").Range("D249").Value = CType(bdrow.bolt_diameter, Double)
            '                            End If
            '                            '.Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(bdrow.bolt_material, String)
            '                            If Not IsNothing(bdrow.bolt_thread_type) Then
            '                                .Worksheets("Main").Range("D254").Value = CType(bdrow.bolt_thread_type, String)
            '                            End If
            '                            'If Not IsNothing(bdrow.bolt_circle) Then
            '                            '    .Worksheets("Main").Range("P" & PlateRow2).Value = CType(bdrow.bolt_circle, Double)
            '                            'End If

            '                            If Not IsNothing(bdrow.eta_factor) Then
            '                                .Worksheets("Main").Range("D253").Value = CType(bdrow.eta_factor, Double)
            '                            End If
            '                            If Not IsNothing(bdrow.lar) Then
            '                                .Worksheets("Main").Range("D252").Value = CType(bdrow.lar, Double)
            '                            End If

            '                            'Need to store elevation also within Database Tab 
            '                            If Not IsNothing(row.connection_elevation) Then
            '                                .Worksheets("Database").Cells(myrow2 - 24, mycol).Value = CType(row.connection_elevation, Double)
            '                            End If
            '                        ElseIf structure_type = "Monopole" Then
            '                            If Not IsNothing(bgrow.BoltDetails.Count) Then
            '                                .Worksheets("MP Connection Summary").Range("L" & PlateRow2).Value = CType(bgrow.BoltDetails.Count, Integer)
            '                            End If
            '                            If Not IsNothing(bdrow.bolt_diameter) Then
            '                                .Worksheets("MP Connection Summary").Range("M" & PlateRow2).Value = CType(bdrow.bolt_diameter, Double)
            '                            End If
            '                            '.Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(bdrow.bolt_material, String)
            '                            If Not IsNothing(bdrow.bolt_thread_type) Then
            '                                .Worksheets("MP Connection Summary").Range("O" & PlateRow2).Value = CType(bdrow.bolt_thread_type, String)
            '                            End If
            '                            If Not IsNothing(bdrow.bolt_circle) Then
            '                                .Worksheets("MP Connection Summary").Range("P" & PlateRow2).Value = CType(bdrow.bolt_circle, Double)
            '                            End If
            '                            If row.connection_type = "Base" Then
            '                                If Not IsNothing(bdrow.eta_factor) Then
            '                                    .Worksheets("MP Connection Summary").Range("N11").Value = CType(bdrow.eta_factor, Double)
            '                                End If
            '                                If Not IsNothing(bdrow.lar) Then
            '                                    .Worksheets("MP Connection Summary").Range("N10").Value = CType(bdrow.lar, Double)
            '                                End If
            '                            End If
            '                            'Need to store elevation also within Database Tab 
            '                            If Not IsNothing(row.connection_elevation) Then
            '                                .Worksheets("Database").Cells(myrow2 - 24, mycol).Value = CType(row.connection_elevation, Double)
            '                            End If
            '                        End If
            '                    ElseIf FirstRow And row.bolt_configuration = "Custom" Then
            '                        If structure_type = "Self Support" Then
            '                            .Worksheets("Main").Range("D253").ClearContents
            '                            .Worksheets("Main").Range("D252").ClearContents
            '                            .Worksheets("Main").Range("D254").ClearContents
            '                        End If
            '                    End If


            '                    If Not IsNothing(bdrow.ID) Then
            '                        .Worksheets("Sub Tables (SAPI)").Range("AG" & BoltDRow + bump2).Value = CType(bdrow.ID, Integer)
            '                    End If
            '                    If Not IsNothing(bdrow.bolt_group_id) Then
            '                        .Worksheets("Sub Tables (SAPI)").Range("AH" & BoltDRow + bump2).Value = CType(bdrow.bolt_group_id, Integer)
            '                    End If

            '                    'If Not IsNothing(bdrow.bolt_id) Then
            '                    '    .Worksheets("Database").Cells(myrow, mycol).Value = CType(bdrow.bolt_id, Integer)
            '                    'End If
            '                    .Worksheets("Database").Cells(myrow2, mycol).Value = CType(bgid, Integer)
            '                    If Not IsNothing(bdrow.bolt_location) Then
            '                        .Worksheets("Database").Cells(myrow2 + 4, mycol).Value = CType(bdrow.bolt_location, Double)
            '                    End If
            '                    If Not IsNothing(bdrow.bolt_diameter) Then
            '                        .Worksheets("Database").Cells(myrow2 + 1, mycol).Value = CType(bdrow.bolt_diameter, Double)
            '                    End If
            '                    'If Not IsNothing(bdrow.bolt_material) Then
            '                    '    .Worksheets("Database").Cells(myrow2 + 2, mycol).Value = CType(bdrow.bolt_material, Integer)
            '                    'End If

            '                    For Each mrow As CCIplateMaterial In bdrow.CCIplateMaterials

            '                        If mrow.default_material = True Then
            '                            If FirstRow And row.bolt_configuration = "Symmetrical" Then
            '                                If structure_type = "Self Support" Then
            '                                    If Not IsNothing(mrow.name) Then
            '                                        .Worksheets("Main").Range("D250").Value = CType(mrow.name, String)
            '                                    End If
            '                                ElseIf structure_type = "Monopole" Then
            '                                    If Not IsNothing(mrow.name) Then
            '                                        .Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(mrow.name, String)
            '                                    End If
            '                                End If
            '                            End If
            '                            If Not IsNothing(mrow.name) Then
            '                                .Worksheets("Database").Cells(myrow2 + 2, mycol).Value = CType(mrow.name, String)
            '                            End If
            '                        Else 'After adding new materail, save material name in a list to reference for other plates to see if materail was already added. 
            '                            For Each tmrow In tempMaterials
            '                                If mrow.ID = tmrow.ID Then
            '                                    matflag = True 'don't add to excel
            '                                    Exit For
            '                                End If
            '                            Next
            '                            If matflag = False Then
            '                                tempMaterial = New CCIplateMaterial(mrow.ID)
            '                                tempMaterials.Add(tempMaterial)

            '                                ''.Worksheets("Sub Tables (SAPI)").Range("AR").Count
            '                                'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Columns("AR").Count 'counts total columns in excel
            '                                'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Cells("AR:AR").Count
            '                                'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").GetDataRange("AR:AR")

            '                                If FirstRow And row.bolt_configuration = "Symmetrical" Then
            '                                    If structure_type = "Self Support" Then
            '                                        If Not IsNothing(mrow.name) Then
            '                                            .Worksheets("Main").Range("D250").Value = CType(mrow.name, String)
            '                                        End If
            '                                    ElseIf structure_type = "Monopole" Then
            '                                        If Not IsNothing(mrow.name) Then
            '                                            .Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(mrow.name, String)
            '                                        End If
            '                                    End If
            '                                End If
            '                                If Not IsNothing(mrow.name) Then
            '                                    .Worksheets("Database").Cells(myrow2 + 2, mycol).Value = CType(mrow.name, String)
            '                                End If
            '                                'planning to reference SaveMaterial here since variables will be similar between sources
            '                                SaveMaterial(wb, mrow, MatRow)
            '                                MatRow += 1
            '                            Else
            '                                'If Not IsNothing(mrow.name) Then
            '                                '    .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).Value = CType(mrow.name, String)
            '                                'Else
            '                                '    .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).ClearContents
            '                                'End If

            '                                If FirstRow And row.bolt_configuration = "Symmetrical" Then
            '                                    If structure_type = "Self Support" Then
            '                                        If Not IsNothing(mrow.name) Then
            '                                            .Worksheets("Main").Range("D250").Value = CType(mrow.name, String)
            '                                        End If
            '                                    ElseIf structure_type = "Monopole" Then
            '                                        If Not IsNothing(mrow.name) Then
            '                                            .Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(mrow.name, String)
            '                                        End If
            '                                    End If
            '                                End If
            '                                If Not IsNothing(mrow.name) Then
            '                                    .Worksheets("Database").Cells(myrow2 + 2, mycol).Value = CType(mrow.name, String)
            '                                End If

            '                            End If
            '                        End If
            '                        matflag = False 'reset material flag
            '                    Next

            '                    If Not IsNothing(bdrow.bolt_circle) Then
            '                        .Worksheets("Database").Cells(myrow2 + 3, mycol).Value = CType(bdrow.bolt_circle, Double)
            '                    End If
            '                    If Not IsNothing(bdrow.eta_factor) Then
            '                        .Worksheets("Database").Cells(myrow2 + 5, mycol).Value = CType(bdrow.eta_factor, Double)
            '                    End If
            '                    If Not IsNothing(bdrow.lar) Then
            '                        .Worksheets("Database").Cells(myrow2 + 6, mycol).Value = CType(bdrow.lar, Double)
            '                    End If
            '                    If Not IsNothing(bdrow.bolt_thread_type) Then
            '                        .Worksheets("Database").Cells(myrow2 + 7, mycol).Value = CType(bdrow.bolt_thread_type, String)
            '                    End If
            '                    If Not IsNothing(bdrow.area_override) Then
            '                        .Worksheets("Database").Cells(myrow2 + 8, mycol).Value = CType(bdrow.area_override, Double)
            '                    End If
            '                    If Not IsNothing(bdrow.tension_only) Then
            '                        If bdrow.tension_only = True Then
            '                            .Worksheets("Database").Cells(myrow2 + 9, mycol).Value = "Yes"
            '                        Else
            '                            .Worksheets("Database").Cells(myrow2 + 9, mycol).Value = "No"
            '                        End If
            '                    End If
            '                    myrow2 += 10
            '                    BoltDRow += 1
            '                    FirstRow = False 'Turns off saving bolt information to MP Connection Summary Tab if not first row
            '                End If
            '            Next

            '            myrow += 4
            '            myrow3 += 1
            '            bgid += 1
            '            BoltGRow += 1

            '        End If

            '    Next

            '    For Each bsdrow As BridgeStiffenerDetail In row.BridgeStiffenerDetails
            '        If bsdrow.connection_id = row.ID Then

            '            If Not IsNothing(bsdrow.ID) Then
            '                .Worksheets("Sub Tables (SAPI)").Range("CF" & BridgeDRow).Value = CType(bsdrow.ID, Integer)
            '                .Worksheets("Sub Tables (SAPI)").Range("CG" & BridgeDRow).Value = CType(row.ID, Integer) 'connetion id req. when deleting
            '            End If
            '            If Not IsNothing(bsdrow.connection_id) Then
            '                '.Worksheets("MP Connection Summary").Range("B" & BridgeDRow + 164).Value = CType(bsdrow.plate_id, Integer)
            '                .Worksheets("MP Connection Summary").Range("B" & BridgeDRow + 164).Value = CType(row.connection_elevation, Integer)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("B" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.stiffener_type) Then
            '                .Worksheets("MP Connection Summary").Range("C" & BridgeDRow + 164).Value = CType(bsdrow.stiffener_type, String)
            '            End If
            '            If Not IsNothing(bsdrow.analysis_type) Then
            '                .Worksheets("MP Connection Summary").Range("D" & BridgeDRow + 164).Value = CType(bsdrow.analysis_type, String)
            '            End If
            '            If Not IsNothing(bsdrow.quantity) Then
            '                .Worksheets("MP Connection Summary").Range("E" & BridgeDRow + 164).Value = CType(bsdrow.quantity, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("E" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.bridge_stiffener_width) Then
            '                .Worksheets("MP Connection Summary").Range("F" & BridgeDRow + 164).Value = CType(bsdrow.bridge_stiffener_width, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("F" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.bridge_stiffener_thickness) Then
            '                .Worksheets("MP Connection Summary").Range("G" & BridgeDRow + 164).Value = CType(bsdrow.bridge_stiffener_thickness, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("G" & BridgeDRow + 164).ClearContents
            '            End If
            '            'If Not IsNothing(bsdrow.bridge_stiffener_material) Then
            '            '    .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).Value = CType(bsdrow.bridge_stiffener_material, Integer)
            '            'Else
            '            '    .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).ClearContents
            '            'End If
            '            For Each mrow As CCIplateMaterial In bsdrow.CCIplateMaterials
            '                If mrow.default_material = True Then
            '                    If Not IsNothing(mrow.name) Then
            '                        .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).Value = CType(mrow.name, String)
            '                    Else
            '                        .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).ClearContents
            '                    End If
            '                Else 'After adding new materail, save material name in a list to reference for other plates to see if materail was already added. 
            '                    For Each tmrow In tempMaterials
            '                        If mrow.ID = tmrow.ID Then
            '                            matflag = True 'don't add to excel
            '                            Exit For
            '                        End If
            '                    Next
            '                    If matflag = False Then
            '                        tempMaterial = New CCIplateMaterial(mrow.ID)
            '                        tempMaterials.Add(tempMaterial)

            '                        '.Worksheets("Sub Tables (SAPI)").Range("AR").Count
            '                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Columns("AR").Count 'counts total columns in excel
            '                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Cells("AR:AR").Count
            '                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").GetDataRange("AR:AR")

            '                        'If Not IsNothing(mrow.ID) Then
            '                        '    .Worksheets("Sub Tables (SAPI)").Range("AR" & MatRow2).Value = CType(mrow.ID, Integer)
            '                        'End If
            '                        If Not IsNothing(mrow.name) Then
            '                            .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).Value = CType(mrow.name, String)
            '                        Else
            '                            .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).ClearContents
            '                        End If

            '                        SaveMaterial(wb, mrow, MatRow)
            '                        MatRow += 1
            '                    Else
            '                        If Not IsNothing(mrow.name) Then
            '                            .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).Value = CType(mrow.name, String)
            '                        Else
            '                            .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).ClearContents
            '                        End If
            '                    End If
            '                End If
            '                matflag = False 'reset flag
            '            Next
            '            If Not IsNothing(bsdrow.unbraced_length) Then
            '                .Worksheets("MP Connection Summary").Range("I" & BridgeDRow + 164).Value = CType(bsdrow.unbraced_length, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("I" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.total_length) Then
            '                .Worksheets("MP Connection Summary").Range("J" & BridgeDRow + 164).Value = CType(bsdrow.total_length, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("J" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.weld_size) Then
            '                .Worksheets("MP Connection Summary").Range("K" & BridgeDRow + 164).Value = CType(bsdrow.weld_size, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("K" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.exx) Then
            '                .Worksheets("MP Connection Summary").Range("L" & BridgeDRow + 164).Value = CType(bsdrow.exx, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("L" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.upper_weld_length) Then
            '                .Worksheets("MP Connection Summary").Range("M" & BridgeDRow + 164).Value = CType(bsdrow.upper_weld_length, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("M" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.lower_weld_length) Then
            '                .Worksheets("MP Connection Summary").Range("N" & BridgeDRow + 164).Value = CType(bsdrow.lower_weld_length, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("N" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.upper_plate_width) Then
            '                .Worksheets("MP Connection Summary").Range("O" & BridgeDRow + 164).Value = CType(bsdrow.upper_plate_width, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("O" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.lower_plate_width) Then
            '                .Worksheets("MP Connection Summary").Range("P" & BridgeDRow + 164).Value = CType(bsdrow.lower_plate_width, Double)
            '            Else
            '                .Worksheets("MP Connection Summary").Range("P" & BridgeDRow + 164).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.neglect_flange_connection) Then
            '                If bsdrow.neglect_flange_connection = True Then
            '                    .Worksheets("MP Connection Summary").Range("R" & BridgeDRow + 164).Value = "Yes"
            '                Else
            '                    .Worksheets("MP Connection Summary").Range("R" & BridgeDRow + 164).Value = "No"
            '                End If
            '            End If
            '            If Not IsNothing(bsdrow.bolt_hole_diameter) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("AQ" & BridgeDRow + 25).Value = CType(bsdrow.bolt_hole_diameter, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("AQ" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.bolt_qty_eccentric) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("DE" & BridgeDRow + 25).Value = CType(bsdrow.bolt_qty_eccentric, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("DE" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.bolt_qty_shear) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("DF" & BridgeDRow + 25).Value = CType(bsdrow.bolt_qty_shear, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("DF" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.intermediate_bolt_spacing) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("DG" & BridgeDRow + 25).Value = CType(bsdrow.intermediate_bolt_spacing, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("DG" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.bolt_diameter) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("DH" & BridgeDRow + 25).Value = CType(bsdrow.bolt_diameter, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("DH" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.bolt_sleeve_diameter) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("DJ" & BridgeDRow + 25).Value = CType(bsdrow.bolt_sleeve_diameter, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("DJ" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.washer_diameter) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("DL" & BridgeDRow + 25).Value = CType(bsdrow.washer_diameter, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("DL" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.bolt_tensile_strength) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("DN" & BridgeDRow + 25).Value = CType(bsdrow.bolt_tensile_strength, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("DN" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.bolt_allowable_shear) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("DP" & BridgeDRow + 25).Value = CType(bsdrow.bolt_allowable_shear, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("DP" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.exx_shim_plate) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("ES" & BridgeDRow + 25).Value = CType(bsdrow.exx_shim_plate, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("ES" & BridgeDRow + 25).ClearContents
            '            End If
            '            If Not IsNothing(bsdrow.filler_shim_thickness) Then
            '                .Worksheets("Bridge Stiffener Calcs").Range("ET" & BridgeDRow + 25).Value = CType(bsdrow.filler_shim_thickness, Double)
            '            Else
            '                .Worksheets("Bridge Stiffener Calcs").Range("ET" & BridgeDRow + 25).ClearContents
            '            End If

            '            BridgeDRow += 1
            '            'BridgeDRow2 += 1

            '        End If
            '    Next

            '    PlateRow += 1
            '    PlateRow2 -= 2
            '    'PlateRow3 -= 1
            '    mycol += 1
            '    bump += 5
            '    bump2 += 100

            'Next

            ''Update named ranges so dropdown seletions work correctly
            ''Materials
            'Dim qty As Integer
            'qty = tempMaterials.Count
            ''Dim definedName As DefinedName = .DefinedNames.Add("Materials", "Materials!$B$5:$D$" & qty + 39)
            ''Dim definedName As DefinedName = .DefinedNames.
            '''Dim definedName As DefinedName = .DefinedNames.scope()
            ''Dim rangeC2D3 As CellRange = .Range(definedName.Name)
            '''IWorkbook.DefinedNames

            'Dim definedName As DefinedName = .DefinedNames.GetDefinedName("Materials")
            'definedName.RefersTo = "Materials!$B$5:$B$" & qty + 39


            'End If







            ''Hiding/unhiding specific tabs
            'If Me.pile_group_config = "Circular" Then
            '    .Worksheets("Moment of Inertia").Visible = False
            '    .Worksheets("Moment of Inertia (Circle)").Visible = True
            'Else
            '    .Worksheets("Moment of Inertia").Visible = True
            '    .Worksheets("Moment of Inertia (Circle)").Visible = False
            'End If

            ''Resizing Image 'User is currently running solution which will resize image within tool
            ''Try
            ''    With .Worksheets("Input").Charts(0)
            ''        .Width = (300 / Math.Max(CType(pf.pad_width_dir1, Double), CType(pf.pad_width_dir2, Double))) * CType(pf.pad_width_dir1, Double) * 4.19 '4.19 multiplier determined through trial and error. 
            ''        .Height = (300 / Math.Max(CType(pf.pad_width_dir1, Double), CType(pf.pad_width_dir2, Double))) * CType(pf.pad_width_dir2, Double) * 4.19
            ''    End With
            ''Catch
            ''    'error handling to avoid dividing by zero
            ''End Try


        End With

    End Sub

#End Region

#Region "Save to EDS"

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tool_version.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Structural_105.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Structural_105")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tool_version = " & Me.tool_version.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit = " & Me.bus_unit.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id = " & Me.structure_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Structural_105 = " & Me.Structural_105.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function

#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As LegReinforcement = TryCast(other, LegReinforcement)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.tool_version.CheckChange(otherToCompare.tool_version, changes, categoryName, "Tool Version"), Equals, False)
        'Equals = If(Me.bus_unit.CheckChange(otherToCompare.bus_unit, changes, categoryName, "Bus Unit"), Equals, False)
        'Equals = If(Me.structure_id.CheckChange(otherToCompare.structure_id, changes, categoryName, "Structure Id"), Equals, False)
        'Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        'Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)
        Equals = If(Me.Structural_105.CheckChange(otherToCompare.Structural_105, changes, categoryName, "Structural 105"), Equals, False)

        'Details
        If Me.LegReinforcementDetails.Count > 0 Then
            Equals = If(Me.LegReinforcementDetails.CheckChange(otherToCompare.LegReinforcementDetails, changes, categoryName, "Leg Reinforcement Details"), Equals, False)
        End If

        Return Equals

    End Function
#End Region

End Class

Partial Public Class LegReinforcementDetail
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "Leg Reinforcement Details"
    Public Overrides ReadOnly Property EDSTableName As String = "tnx.memb_leg_reinforcement_details"

    Public Overrides Function SQLInsert() As String

        SQLInsert = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Detail_INSERT
        SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT DETAIL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT DETAIL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        ''Connection Results
        'For Each row As ConnectionResults In ConnectionResults
        '    SQLInsert = SQLInsert.Replace("--BEGIN --[CONNECTION RESULTS INSERT BEGIN]", "BEGIN --[CONNECTION RESULTS INSERT BEGIN]")
        '    SQLInsert = SQLInsert.Replace("--END --[CONNECTION RESULTS INSERT END]", "END --[CONNECTION RESULTS INSERT END]")
        '    SQLInsert = SQLInsert.Replace("--[CONNECTION RESULTS INSERT]", row.SQLInsert)
        'Next

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        SQLUpdate = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Detail_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        ''Connection Results
        'For Each row As ConnectionResults In ConnectionResults
        '    SQLUpdate = SQLUpdate.Replace("--BEGIN --[CONNECTION RESULTS INSERT BEGIN]", "BEGIN --[CONNECTION RESULTS INSERT BEGIN]")
        '    SQLUpdate = SQLUpdate.Replace("--END --[CONNECTION RESULTS INSERT END]", "END --[CONNECTION RESULTS INSERT END]")
        '    SQLUpdate = SQLUpdate.Replace("--[CONNECTION RESULTS INSERT]", row.SQLInsert)
        'Next

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        SQLDelete = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Detail_DELETE
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    'Private _local_id As Integer?
    Private _ID As Integer?
    Private _leg_reinforcement_id As Integer?
    Private _leg_load_time_mod_option As Boolean?
    Private _end_connection_type As String
    Private _leg_crushing As Boolean?
    Private _applied_load_type As String
    Private _slenderness_ratio_type As String
    Private _intermeditate_connection_type As String
    Private _intermeditate_connection_spacing As Double?
    Private _ki_override As Double?
    Private _leg_diameter As Double?
    Private _leg_thickness As Double?
    Private _leg_grade As Double?
    Private _leg_unbraced_length As Double?
    Private _rein_diameter As Double?
    Private _rein_thickness As Double?
    Private _rein_grade As Double?
    Private _print_bolt_on_connections As Boolean?
    Private _leg_length As Double?
    Private _rein_length As Double?
    Private _set_top_to_bottom As Boolean?
    Private _flange_bolt_quantity_bot As Integer?
    Private _flange_bolt_circle_bot As Double?
    Private _flange_bolt_orientation_bot As Integer?
    Private _flange_bolt_quantity_top As Integer?
    Private _flange_bolt_circle_top As Double?
    Private _flange_bolt_orientation_top As Integer?
    Private _threaded_rod_size_bot As String
    Private _threaded_rod_mat_bot As String
    Private _threaded_rod_quantity_bot As Integer?
    Private _threaded_rod_unbraced_length_bot As Double?
    Private _threaded_rod_size_top As String
    Private _threaded_rod_mat_top As String
    Private _threaded_rod_quantity_top As Integer?
    Private _threaded_rod_unbraced_length_top As Double?
    Private _stiffener_height_bot As Double?
    Private _stiffener_length_bot As Double?
    Private _stiffener_fillet_bot As Integer?
    Private _stiffener_exx_bot As Double?
    Private _flange_thickness_bot As Double?
    Private _stiffener_height_top As Double?
    Private _stiffener_length_top As Double?
    Private _stiffener_fillet_top As Integer?
    Private _stiffener_exx_top As Double?
    Private _flange_thickness_top As Double?
    Private _structure_ind As String
    Private _reinforcement_type As String
    Private _leg_reinforcement_name As String
    Private _local_id As Integer?
    Private _top_elev As Double?
    Private _bot_elev As Double?


    'Public Property ConnectionResults As New List(Of ConnectionResults)

    '<Category("Leg Reinforcement Details"), Description(""), DisplayName("Local Id")>
    'Public Property local_id() As Integer?
    '    Get
    '        Return Me._local_id
    '    End Get
    '    Set
    '        Me._local_id = Value
    '    End Set
    'End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Id")>
    Public Property leg_reinforcement_id() As Integer?
        Get
            Return Me._leg_reinforcement_id
        End Get
        Set
            Me._leg_reinforcement_id = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Load Time Mod Option")>
    Public Property leg_load_time_mod_option() As Boolean?
        Get
            Return Me._leg_load_time_mod_option
        End Get
        Set
            Me._leg_load_time_mod_option = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("End Connection Type")>
    Public Property end_connection_type() As String
        Get
            Return Me._end_connection_type
        End Get
        Set
            Me._end_connection_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Crushing")>
    Public Property leg_crushing() As Boolean?
        Get
            Return Me._leg_crushing
        End Get
        Set
            Me._leg_crushing = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Applied Load Type")>
    Public Property applied_load_type() As String
        Get
            Return Me._applied_load_type
        End Get
        Set
            Me._applied_load_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Slenderness Ratio Type")>
    Public Property slenderness_ratio_type() As String
        Get
            Return Me._slenderness_ratio_type
        End Get
        Set
            Me._slenderness_ratio_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Intermeditate Connection Type")>
    Public Property intermeditate_connection_type() As String
        Get
            Return Me._intermeditate_connection_type
        End Get
        Set
            Me._intermeditate_connection_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Intermeditate Connection Spacing")>
    Public Property intermeditate_connection_spacing() As Double?
        Get
            Return Me._intermeditate_connection_spacing
        End Get
        Set
            Me._intermeditate_connection_spacing = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Ki Override")>
    Public Property ki_override() As Double?
        Get
            Return Me._ki_override
        End Get
        Set
            Me._ki_override = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Diameter")>
    Public Property leg_diameter() As Double?
        Get
            Return Me._leg_diameter
        End Get
        Set
            Me._leg_diameter = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Thickness")>
    Public Property leg_thickness() As Double?
        Get
            Return Me._leg_thickness
        End Get
        Set
            Me._leg_thickness = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Grade")>
    Public Property leg_grade() As Double?
        Get
            Return Me._leg_grade
        End Get
        Set
            Me._leg_grade = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Unbraced Length")>
    Public Property leg_unbraced_length() As Double?
        Get
            Return Me._leg_unbraced_length
        End Get
        Set
            Me._leg_unbraced_length = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Rein Diameter")>
    Public Property rein_diameter() As Double?
        Get
            Return Me._rein_diameter
        End Get
        Set
            Me._rein_diameter = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Rein Thickness")>
    Public Property rein_thickness() As Double?
        Get
            Return Me._rein_thickness
        End Get
        Set
            Me._rein_thickness = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Rein Grade")>
    Public Property rein_grade() As Double?
        Get
            Return Me._rein_grade
        End Get
        Set
            Me._rein_grade = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Print Bolt On Connections")>
    Public Property print_bolt_on_connections() As Boolean?
        Get
            Return Me._print_bolt_on_connections
        End Get
        Set
            Me._print_bolt_on_connections = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Length")>
    Public Property leg_length() As Double?
        Get
            Return Me._leg_length
        End Get
        Set
            Me._leg_length = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Rein Length")>
    Public Property rein_length() As Double?
        Get
            Return Me._rein_length
        End Get
        Set
            Me._rein_length = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Set Top To Bottom")>
    Public Property set_top_to_bottom() As Boolean?
        Get
            Return Me._set_top_to_bottom
        End Get
        Set
            Me._set_top_to_bottom = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Quantity Bot")>
    Public Property flange_bolt_quantity_bot() As Integer?
        Get
            Return Me._flange_bolt_quantity_bot
        End Get
        Set
            Me._flange_bolt_quantity_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Circle Bot")>
    Public Property flange_bolt_circle_bot() As Double?
        Get
            Return Me._flange_bolt_circle_bot
        End Get
        Set
            Me._flange_bolt_circle_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Orientation Bot")>
    Public Property flange_bolt_orientation_bot() As Integer?
        Get
            Return Me._flange_bolt_orientation_bot
        End Get
        Set
            Me._flange_bolt_orientation_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Quantity Top")>
    Public Property flange_bolt_quantity_top() As Integer?
        Get
            Return Me._flange_bolt_quantity_top
        End Get
        Set
            Me._flange_bolt_quantity_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Circle Top")>
    Public Property flange_bolt_circle_top() As Double?
        Get
            Return Me._flange_bolt_circle_top
        End Get
        Set
            Me._flange_bolt_circle_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Orientation Top")>
    Public Property flange_bolt_orientation_top() As Integer?
        Get
            Return Me._flange_bolt_orientation_top
        End Get
        Set
            Me._flange_bolt_orientation_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Size Bot")>
    Public Property threaded_rod_size_bot() As String
        Get
            Return Me._threaded_rod_size_bot
        End Get
        Set
            Me._threaded_rod_size_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Mat Bot")>
    Public Property threaded_rod_mat_bot() As String
        Get
            Return Me._threaded_rod_mat_bot
        End Get
        Set
            Me._threaded_rod_mat_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Quantity Bot")>
    Public Property threaded_rod_quantity_bot() As Integer?
        Get
            Return Me._threaded_rod_quantity_bot
        End Get
        Set
            Me._threaded_rod_quantity_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Unbraced Length Bot")>
    Public Property threaded_rod_unbraced_length_bot() As Double?
        Get
            Return Me._threaded_rod_unbraced_length_bot
        End Get
        Set
            Me._threaded_rod_unbraced_length_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Size Top")>
    Public Property threaded_rod_size_top() As String
        Get
            Return Me._threaded_rod_size_top
        End Get
        Set
            Me._threaded_rod_size_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Mat Top")>
    Public Property threaded_rod_mat_top() As String
        Get
            Return Me._threaded_rod_mat_top
        End Get
        Set
            Me._threaded_rod_mat_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Quantity Top")>
    Public Property threaded_rod_quantity_top() As Integer?
        Get
            Return Me._threaded_rod_quantity_top
        End Get
        Set
            Me._threaded_rod_quantity_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Unbraced Length Top")>
    Public Property threaded_rod_unbraced_length_top() As Double?
        Get
            Return Me._threaded_rod_unbraced_length_top
        End Get
        Set
            Me._threaded_rod_unbraced_length_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Height Bot")>
    Public Property stiffener_height_bot() As Double?
        Get
            Return Me._stiffener_height_bot
        End Get
        Set
            Me._stiffener_height_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Length Bot")>
    Public Property stiffener_length_bot() As Double?
        Get
            Return Me._stiffener_length_bot
        End Get
        Set
            Me._stiffener_length_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Fillet Bot")>
    Public Property stiffener_fillet_bot() As Integer?
        Get
            Return Me._stiffener_fillet_bot
        End Get
        Set
            Me._stiffener_fillet_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Exx Bot")>
    Public Property stiffener_exx_bot() As Double?
        Get
            Return Me._stiffener_exx_bot
        End Get
        Set
            Me._stiffener_exx_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Thickness Bot")>
    Public Property flange_thickness_bot() As Double?
        Get
            Return Me._flange_thickness_bot
        End Get
        Set
            Me._flange_thickness_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Height Top")>
    Public Property stiffener_height_top() As Double?
        Get
            Return Me._stiffener_height_top
        End Get
        Set
            Me._stiffener_height_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Length Top")>
    Public Property stiffener_length_top() As Double?
        Get
            Return Me._stiffener_length_top
        End Get
        Set
            Me._stiffener_length_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Fillet Top")>
    Public Property stiffener_fillet_top() As Integer?
        Get
            Return Me._stiffener_fillet_top
        End Get
        Set
            Me._stiffener_fillet_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Exx Top")>
    Public Property stiffener_exx_top() As Double?
        Get
            Return Me._stiffener_exx_top
        End Get
        Set
            Me._stiffener_exx_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Thickness Top")>
    Public Property flange_thickness_top() As Double?
        Get
            Return Me._flange_thickness_top
        End Get
        Set
            Me._flange_thickness_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Structure Ind")>
    Public Property structure_ind() As String
        Get
            Return Me._structure_ind
        End Get
        Set
            Me._structure_ind = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Reinforcement Type")>
    Public Property reinforcement_type() As String
        Get
            Return Me._reinforcement_type
        End Get
        Set
            Me._reinforcement_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Name")>
    Public Property leg_reinforcement_name() As String
        Get
            Return Me._leg_reinforcement_name
        End Get
        Set
            Me._leg_reinforcement_name = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Local Id")>
    Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Top Elev")>
    Public Property top_elev() As Double?
        Get
            Return Me._top_elev
        End Get
        Set
            Me._top_elev = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Bot Elev")>
    Public Property bot_elev() As Double?
        Get
            Return Me._bot_elev
        End Get
        Set
            Me._bot_elev = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal row As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal prow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = row
        'probably need a local id in order to pull correct results
        'If EDStruefalse = False Then 'Only pull in local id when referencing Excel
        '    Me.local_id = DBtoNullableInt(dr.Item("local_connection_id"))
        'End If

        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = True Then 'Only Pull in when referencing EDS
            Me.leg_reinforcement_id = DBtoNullableInt(dr.Item("leg_reinforcement_id"))
        End If
        Me.leg_load_time_mod_option = If(EDStruefalse, DBtoNullableBool(dr.Item("leg_load_time_mod_option")), If(DBtoStr(dr.Item("leg_load_time_mod_option")) = "Yes", True, If(DBtoStr(dr.Item("leg_load_time_mod_option")) = "No", False, DBtoNullableBool(dr.Item("leg_load_time_mod_option"))))) 'Listed as a string and need to convert to Boolean
        Me.end_connection_type = DBtoStr(dr.Item("end_connection_type"))
        Me.leg_crushing = If(EDStruefalse, DBtoNullableBool(dr.Item("leg_crushing")), If(DBtoStr(dr.Item("leg_crushing")) = "Yes", True, If(DBtoStr(dr.Item("leg_crushing")) = "No", False, DBtoNullableBool(dr.Item("leg_crushing"))))) 'Listed as a string and need to convert to Boolean
        Me.applied_load_type = DBtoStr(dr.Item("applied_load_type"))
        Me.slenderness_ratio_type = DBtoStr(dr.Item("slenderness_ratio_type"))
        Me.intermeditate_connection_type = DBtoStr(dr.Item("intermeditate_connection_type"))
        Me.intermeditate_connection_spacing = DBtoNullableDbl(dr.Item("intermeditate_connection_spacing"))
        Me.ki_override = DBtoNullableDbl(dr.Item("ki_override"))
        Me.leg_diameter = DBtoNullableDbl(dr.Item("leg_diameter"))
        Me.leg_thickness = DBtoNullableDbl(dr.Item("leg_thickness"))
        Me.leg_grade = DBtoNullableDbl(dr.Item("leg_grade"))
        Me.leg_unbraced_length = DBtoNullableDbl(dr.Item("leg_unbraced_length"))
        Me.rein_diameter = DBtoNullableDbl(dr.Item("rein_diameter"))
        Me.rein_thickness = DBtoNullableDbl(dr.Item("rein_thickness"))
        Me.rein_grade = DBtoNullableDbl(dr.Item("rein_grade"))
        Me.print_bolt_on_connections = DBtoNullableBool(dr.Item("print_bolt_on_connections"))
        Me.leg_length = DBtoNullableDbl(dr.Item("leg_length"))
        Me.rein_length = DBtoNullableDbl(dr.Item("rein_length"))
        Me.set_top_to_bottom = DBtoNullableBool(dr.Item("set_top_to_bottom"))
        Me.flange_bolt_quantity_bot = DBtoNullableInt(dr.Item("flange_bolt_quantity_bot"))
        Me.flange_bolt_circle_bot = DBtoNullableDbl(dr.Item("flange_bolt_circle_bot"))
        Me.flange_bolt_orientation_bot = DBtoNullableInt(dr.Item("flange_bolt_orientation_bot"))
        Me.flange_bolt_quantity_top = DBtoNullableInt(dr.Item("flange_bolt_quantity_top"))
        Me.flange_bolt_circle_top = DBtoNullableDbl(dr.Item("flange_bolt_circle_top"))
        Me.flange_bolt_orientation_top = DBtoNullableInt(dr.Item("flange_bolt_orientation_top"))
        Me.threaded_rod_size_bot = DBtoStr(dr.Item("threaded_rod_size_bot"))
        Me.threaded_rod_mat_bot = DBtoStr(dr.Item("threaded_rod_mat_bot"))
        Me.threaded_rod_quantity_bot = DBtoNullableInt(dr.Item("threaded_rod_quantity_bot"))
        Me.threaded_rod_unbraced_length_bot = DBtoNullableDbl(dr.Item("threaded_rod_unbraced_length_bot"))
        Me.threaded_rod_size_top = DBtoStr(dr.Item("threaded_rod_size_top"))
        Me.threaded_rod_mat_top = DBtoStr(dr.Item("threaded_rod_mat_top"))
        Me.threaded_rod_quantity_top = DBtoNullableInt(dr.Item("threaded_rod_quantity_top"))
        Me.threaded_rod_unbraced_length_top = DBtoNullableDbl(dr.Item("threaded_rod_unbraced_length_top"))
        Me.stiffener_height_bot = DBtoNullableDbl(dr.Item("stiffener_height_bot"))
        Me.stiffener_length_bot = DBtoNullableDbl(dr.Item("stiffener_length_bot"))
        Me.stiffener_fillet_bot = DBtoNullableInt(dr.Item("stiffener_fillet_bot"))
        Me.stiffener_exx_bot = DBtoNullableDbl(dr.Item("stiffener_exx_bot"))
        Me.flange_thickness_bot = DBtoNullableDbl(dr.Item("flange_thickness_bot"))
        Me.stiffener_height_top = DBtoNullableDbl(dr.Item("stiffener_height_top"))
        Me.stiffener_length_top = DBtoNullableDbl(dr.Item("stiffener_length_top"))
        Me.stiffener_fillet_top = DBtoNullableInt(dr.Item("stiffener_fillet_top"))
        Me.stiffener_exx_top = DBtoNullableDbl(dr.Item("stiffener_exx_top"))
        Me.flange_thickness_top = DBtoNullableDbl(dr.Item("flange_thickness_top"))
        Me.structure_ind = DBtoStr(dr.Item("structure_ind"))
        Me.reinforcement_type = DBtoStr(dr.Item("reinforcement_type"))
        Me.leg_reinforcement_name = DBtoStr(dr.Item("leg_reinforcement_name"))
        Me.local_id = DBtoNullableInt(dr.Item("local_id"))
        Me.top_elev = DBtoNullableDbl(dr.Item("top_elev"))
        Me.bot_elev = DBtoNullableDbl(dr.Item("bot_elev"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_reinforcement_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_load_time_mod_option.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.end_connection_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_crushing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.applied_load_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.slenderness_ratio_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.intermeditate_connection_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.intermeditate_connection_spacing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ki_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_unbraced_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rein_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rein_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rein_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.print_bolt_on_connections.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rein_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.set_top_to_bottom.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flange_bolt_quantity_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flange_bolt_circle_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flange_bolt_orientation_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flange_bolt_quantity_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flange_bolt_circle_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flange_bolt_orientation_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.threaded_rod_size_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.threaded_rod_mat_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.threaded_rod_quantity_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.threaded_rod_unbraced_length_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.threaded_rod_size_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.threaded_rod_mat_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.threaded_rod_quantity_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.threaded_rod_unbraced_length_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_height_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_length_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_fillet_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_exx_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flange_thickness_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_height_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_length_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_fillet_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_exx_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flange_thickness_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_ind.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reinforcement_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_reinforcement_name.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.top_elev.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bot_elev.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_reinforcement_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_load_time_mod_option")
        SQLInsertFields = SQLInsertFields.AddtoDBString("end_connection_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_crushing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("applied_load_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("slenderness_ratio_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("intermeditate_connection_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("intermeditate_connection_spacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ki_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_unbraced_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rein_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rein_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rein_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("print_bolt_on_connections")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rein_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("set_top_to_bottom")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flange_bolt_quantity_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flange_bolt_circle_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flange_bolt_orientation_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flange_bolt_quantity_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flange_bolt_circle_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flange_bolt_orientation_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("threaded_rod_size_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("threaded_rod_mat_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("threaded_rod_quantity_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("threaded_rod_unbraced_length_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("threaded_rod_size_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("threaded_rod_mat_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("threaded_rod_quantity_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("threaded_rod_unbraced_length_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_height_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_length_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_fillet_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_exx_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flange_thickness_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_height_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_length_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_fillet_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_exx_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flange_thickness_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_ind")
        SQLInsertFields = SQLInsertFields.AddtoDBString("reinforcement_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_reinforcement_name")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("top_elev")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bot_elev")


        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_reinforcement_id = " & "@TopLevelID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_load_time_mod_option = " & Me.leg_load_time_mod_option.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("end_connection_type = " & Me.end_connection_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_crushing = " & Me.leg_crushing.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("applied_load_type = " & Me.applied_load_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("slenderness_ratio_type = " & Me.slenderness_ratio_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("intermeditate_connection_type = " & Me.intermeditate_connection_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("intermeditate_connection_spacing = " & Me.intermeditate_connection_spacing.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ki_override = " & Me.ki_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_diameter = " & Me.leg_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_thickness = " & Me.leg_thickness.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_grade = " & Me.leg_grade.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_unbraced_length = " & Me.leg_unbraced_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rein_diameter = " & Me.rein_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rein_thickness = " & Me.rein_thickness.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rein_grade = " & Me.rein_grade.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("print_bolt_on_connections = " & Me.print_bolt_on_connections.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_length = " & Me.leg_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rein_length = " & Me.rein_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("set_top_to_bottom = " & Me.set_top_to_bottom.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flange_bolt_quantity_bot = " & Me.flange_bolt_quantity_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flange_bolt_circle_bot = " & Me.flange_bolt_circle_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flange_bolt_orientation_bot = " & Me.flange_bolt_orientation_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flange_bolt_quantity_top = " & Me.flange_bolt_quantity_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flange_bolt_circle_top = " & Me.flange_bolt_circle_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flange_bolt_orientation_top = " & Me.flange_bolt_orientation_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("threaded_rod_size_bot = " & Me.threaded_rod_size_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("threaded_rod_mat_bot = " & Me.threaded_rod_mat_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("threaded_rod_quantity_bot = " & Me.threaded_rod_quantity_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("threaded_rod_unbraced_length_bot = " & Me.threaded_rod_unbraced_length_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("threaded_rod_size_top = " & Me.threaded_rod_size_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("threaded_rod_mat_top = " & Me.threaded_rod_mat_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("threaded_rod_quantity_top = " & Me.threaded_rod_quantity_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("threaded_rod_unbraced_length_top = " & Me.threaded_rod_unbraced_length_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_height_bot = " & Me.stiffener_height_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_length_bot = " & Me.stiffener_length_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_fillet_bot = " & Me.stiffener_fillet_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_exx_bot = " & Me.stiffener_exx_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flange_thickness_bot = " & Me.flange_thickness_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_height_top = " & Me.stiffener_height_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_length_top = " & Me.stiffener_length_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_fillet_top = " & Me.stiffener_fillet_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_exx_top = " & Me.stiffener_exx_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flange_thickness_top = " & Me.flange_thickness_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_ind = " & Me.structure_ind.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("reinforcement_type = " & Me.reinforcement_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_reinforcement_name = " & Me.leg_reinforcement_name.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_id = " & Me.local_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("top_elev = " & Me.top_elev.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bot_elev = " & Me.bot_elev.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As LegReinforcementDetail = TryCast(other, LegReinforcementDetail)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.leg_reinforcement_id.CheckChange(otherToCompare.leg_reinforcement_id, changes, categoryName, "Leg Reinforcement Id"), Equals, False)
        Equals = If(Me.leg_load_time_mod_option.CheckChange(otherToCompare.leg_load_time_mod_option, changes, categoryName, "Leg Load Time Mod Option"), Equals, False)
        Equals = If(Me.end_connection_type.CheckChange(otherToCompare.end_connection_type, changes, categoryName, "End Connection Type"), Equals, False)
        Equals = If(Me.leg_crushing.CheckChange(otherToCompare.leg_crushing, changes, categoryName, "Leg Crushing"), Equals, False)
        Equals = If(Me.applied_load_type.CheckChange(otherToCompare.applied_load_type, changes, categoryName, "Applied Load Type"), Equals, False)
        Equals = If(Me.slenderness_ratio_type.CheckChange(otherToCompare.slenderness_ratio_type, changes, categoryName, "Slenderness Ratio Type"), Equals, False)
        Equals = If(Me.intermeditate_connection_type.CheckChange(otherToCompare.intermeditate_connection_type, changes, categoryName, "Intermeditate Connection Type"), Equals, False)
        Equals = If(Me.intermeditate_connection_spacing.CheckChange(otherToCompare.intermeditate_connection_spacing, changes, categoryName, "Intermeditate Connection Spacing"), Equals, False)
        Equals = If(Me.ki_override.CheckChange(otherToCompare.ki_override, changes, categoryName, "Ki Override"), Equals, False)
        Equals = If(Me.leg_diameter.CheckChange(otherToCompare.leg_diameter, changes, categoryName, "Leg Diameter"), Equals, False)
        Equals = If(Me.leg_thickness.CheckChange(otherToCompare.leg_thickness, changes, categoryName, "Leg Thickness"), Equals, False)
        Equals = If(Me.leg_grade.CheckChange(otherToCompare.leg_grade, changes, categoryName, "Leg Grade"), Equals, False)
        Equals = If(Me.leg_unbraced_length.CheckChange(otherToCompare.leg_unbraced_length, changes, categoryName, "Leg Unbraced Length"), Equals, False)
        Equals = If(Me.rein_diameter.CheckChange(otherToCompare.rein_diameter, changes, categoryName, "Rein Diameter"), Equals, False)
        Equals = If(Me.rein_thickness.CheckChange(otherToCompare.rein_thickness, changes, categoryName, "Rein Thickness"), Equals, False)
        Equals = If(Me.rein_grade.CheckChange(otherToCompare.rein_grade, changes, categoryName, "Rein Grade"), Equals, False)
        Equals = If(Me.print_bolt_on_connections.CheckChange(otherToCompare.print_bolt_on_connections, changes, categoryName, "Print Bolt On Connections"), Equals, False)
        Equals = If(Me.leg_length.CheckChange(otherToCompare.leg_length, changes, categoryName, "Leg Length"), Equals, False)
        Equals = If(Me.rein_length.CheckChange(otherToCompare.rein_length, changes, categoryName, "Rein Length"), Equals, False)
        Equals = If(Me.set_top_to_bottom.CheckChange(otherToCompare.set_top_to_bottom, changes, categoryName, "Set Top To Bottom"), Equals, False)
        Equals = If(Me.flange_bolt_quantity_bot.CheckChange(otherToCompare.flange_bolt_quantity_bot, changes, categoryName, "Flange Bolt Quantity Bot"), Equals, False)
        Equals = If(Me.flange_bolt_circle_bot.CheckChange(otherToCompare.flange_bolt_circle_bot, changes, categoryName, "Flange Bolt Circle Bot"), Equals, False)
        Equals = If(Me.flange_bolt_orientation_bot.CheckChange(otherToCompare.flange_bolt_orientation_bot, changes, categoryName, "Flange Bolt Orientation Bot"), Equals, False)
        Equals = If(Me.flange_bolt_quantity_top.CheckChange(otherToCompare.flange_bolt_quantity_top, changes, categoryName, "Flange Bolt Quantity Top"), Equals, False)
        Equals = If(Me.flange_bolt_circle_top.CheckChange(otherToCompare.flange_bolt_circle_top, changes, categoryName, "Flange Bolt Circle Top"), Equals, False)
        Equals = If(Me.flange_bolt_orientation_top.CheckChange(otherToCompare.flange_bolt_orientation_top, changes, categoryName, "Flange Bolt Orientation Top"), Equals, False)
        Equals = If(Me.threaded_rod_size_bot.CheckChange(otherToCompare.threaded_rod_size_bot, changes, categoryName, "Threaded Rod Size Bot"), Equals, False)
        Equals = If(Me.threaded_rod_mat_bot.CheckChange(otherToCompare.threaded_rod_mat_bot, changes, categoryName, "Threaded Rod Mat Bot"), Equals, False)
        Equals = If(Me.threaded_rod_quantity_bot.CheckChange(otherToCompare.threaded_rod_quantity_bot, changes, categoryName, "Threaded Rod Quantity Bot"), Equals, False)
        Equals = If(Me.threaded_rod_unbraced_length_bot.CheckChange(otherToCompare.threaded_rod_unbraced_length_bot, changes, categoryName, "Threaded Rod Unbraced Length Bot"), Equals, False)
        Equals = If(Me.threaded_rod_size_top.CheckChange(otherToCompare.threaded_rod_size_top, changes, categoryName, "Threaded Rod Size Top"), Equals, False)
        Equals = If(Me.threaded_rod_mat_top.CheckChange(otherToCompare.threaded_rod_mat_top, changes, categoryName, "Threaded Rod Mat Top"), Equals, False)
        Equals = If(Me.threaded_rod_quantity_top.CheckChange(otherToCompare.threaded_rod_quantity_top, changes, categoryName, "Threaded Rod Quantity Top"), Equals, False)
        Equals = If(Me.threaded_rod_unbraced_length_top.CheckChange(otherToCompare.threaded_rod_unbraced_length_top, changes, categoryName, "Threaded Rod Unbraced Length Top"), Equals, False)
        Equals = If(Me.stiffener_height_bot.CheckChange(otherToCompare.stiffener_height_bot, changes, categoryName, "Stiffener Height Bot"), Equals, False)
        Equals = If(Me.stiffener_length_bot.CheckChange(otherToCompare.stiffener_length_bot, changes, categoryName, "Stiffener Length Bot"), Equals, False)
        Equals = If(Me.stiffener_fillet_bot.CheckChange(otherToCompare.stiffener_fillet_bot, changes, categoryName, "Stiffener Fillet Bot"), Equals, False)
        Equals = If(Me.stiffener_exx_bot.CheckChange(otherToCompare.stiffener_exx_bot, changes, categoryName, "Stiffener Exx Bot"), Equals, False)
        Equals = If(Me.flange_thickness_bot.CheckChange(otherToCompare.flange_thickness_bot, changes, categoryName, "Flange Thickness Bot"), Equals, False)
        Equals = If(Me.stiffener_height_top.CheckChange(otherToCompare.stiffener_height_top, changes, categoryName, "Stiffener Height Top"), Equals, False)
        Equals = If(Me.stiffener_length_top.CheckChange(otherToCompare.stiffener_length_top, changes, categoryName, "Stiffener Length Top"), Equals, False)
        Equals = If(Me.stiffener_fillet_top.CheckChange(otherToCompare.stiffener_fillet_top, changes, categoryName, "Stiffener Fillet Top"), Equals, False)
        Equals = If(Me.stiffener_exx_top.CheckChange(otherToCompare.stiffener_exx_top, changes, categoryName, "Stiffener Exx Top"), Equals, False)
        Equals = If(Me.flange_thickness_top.CheckChange(otherToCompare.flange_thickness_top, changes, categoryName, "Flange Thickness Top"), Equals, False)
        Equals = If(Me.structure_ind.CheckChange(otherToCompare.structure_ind, changes, categoryName, "Structure Ind"), Equals, False)
        Equals = If(Me.reinforcement_type.CheckChange(otherToCompare.reinforcement_type, changes, categoryName, "Reinforcement Type"), Equals, False)
        Equals = If(Me.leg_reinforcement_name.CheckChange(otherToCompare.leg_reinforcement_name, changes, categoryName, "Leg Reinforcement Name"), Equals, False)
        Equals = If(Me.local_id.CheckChange(otherToCompare.local_id, changes, categoryName, "Local Id"), Equals, False)
        Equals = If(Me.top_elev.CheckChange(otherToCompare.top_elev, changes, categoryName, "Top Elev"), Equals, False)
        Equals = If(Me.bot_elev.CheckChange(otherToCompare.bot_elev, changes, categoryName, "Bot Elev"), Equals, False)

    End Function
#End Region

End Class

'Partial Public Class PlateDetail
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Plate Details"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.plate_details"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Plate Detail (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Plate_Detail_INSERT
'        SQLInsert = SQLInsert.Replace("[PLATE DETAIL VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[PLATE DETAIL FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Plate Material
'        For Each row As CCIplateMaterial In CCIplateMaterials
'            SQLInsert = SQLInsert.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert)
'        Next

'        'Results
'        'If Me.Results.Count > 0 Then
'        For Each row As PlateResults In PlateResults
'            SQLInsert = SQLInsert.Replace("--BEGIN --[PLATE RESULTS INSERT BEGIN]", "BEGIN --[PLATE RESULTS INSERT BEGIN]")
'            SQLInsert = SQLInsert.Replace("--END --[PLATE RESULTS INSERT END]", "END --[PLATE RESULTS INSERT END]")
'            SQLInsert = SQLInsert.Replace("--[PLATE RESULTS INSERT]", row.SQLInsert)
'        Next
'        'End If

'        'Stiffener Group
'        If Me.StiffenerGroups.Count > 0 Then
'            SQLInsert = SQLInsert.Replace("--BEGIN --[STIFFENER GROUP INSERT BEGIN]", "BEGIN --[STIFFENER GROUP INSERT BEGIN]")
'            SQLInsert = SQLInsert.Replace("--END --[STIFFENER GROUP INSERT END]", "END --[STIFFENER GROUP INSERT END]")
'            For Each row As StiffenerGroup In StiffenerGroups
'                SQLInsert = SQLInsert.Replace("--[STIFFENER GROUP INSERT]", row.SQLInsert)
'            Next
'        End If

'        Return SQLInsert

'    End Function

'    Public Overrides Function SQLUpdate() As String

'        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Plate Detail (UPDATE).sql")
'        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Plate_Detail_UPDATE
'        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
'        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Plate Material
'        For Each row As CCIplateMaterial In CCIplateMaterials
'            SQLUpdate = SQLUpdate.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert) 'Can only insert materials, no deleting or updating since database is referenced by all BUs. 
'        Next

'        'Results
'        'If Me.Results.Count > 0 Then
'        '    SQLUpdate = SQLUpdate.Replace("--BEGIN --[RESULTS UPDATE BEGIN]", "BEGIN --[RESULTS UPDATE BEGIN]")
'        '    SQLUpdate = SQLUpdate.Replace("--END --[RESULTS UPDATE END]", "END --[RESULTS UPDATE END]")
'        '    SQLUpdate = SQLUpdate.Replace("--[RESULTS INSERT]", Me.Results.EDSResultQuery)
'        'End If
'        'Insert is always performed for results
'        For Each row As PlateResults In PlateResults
'            SQLUpdate = SQLUpdate.Replace("--BEGIN --[PLATE RESULTS INSERT BEGIN]", "BEGIN --[PLATE RESULTS INSERT BEGIN]")
'            SQLUpdate = SQLUpdate.Replace("--END --[PLATE RESULTS INSERT END]", "END --[PLATE RESULTS INSERT END]")
'            SQLUpdate = SQLUpdate.Replace("--[PLATE RESULTS INSERT]", row.SQLInsert)
'        Next

'        'Stiffener Groups
'        If Me.StiffenerGroups.Count > 0 Then
'            SQLUpdate = SQLUpdate.Replace("--BEGIN --[STIFFENER GROUP UPDATE BEGIN]", "BEGIN --[STIFFENER GROUP UPDATE BEGIN]")
'            SQLUpdate = SQLUpdate.Replace("--END --[STIFFENER GROUP UPDATE END]", "END --[STIFFENER GROUP UPDATE END]")
'            For Each row As StiffenerGroup In StiffenerGroups
'                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
'                    If IsSomethingString(row.stiffener_name) Then
'                        SQLUpdate = SQLUpdate.Replace("--[STIFFENER GROUP INSERT]", row.SQLUpdate)
'                    Else
'                        SQLUpdate = SQLUpdate.Replace("--[STIFFENER GROUP INSERT]", row.SQLDelete)
'                    End If
'                Else
'                    SQLUpdate = SQLUpdate.Replace("--[STIFFENER GROUP INSERT]", row.SQLInsert)
'                End If
'            Next
'        End If

'        Return SQLUpdate

'    End Function

'    Public Overrides Function SQLDelete() As String

'        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Plate Detail (DELETE).sql")
'        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Plate_Detail_DELETE
'        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Stiffener Groups
'        If Me.StiffenerGroups.Count > 0 Then
'            SQLDelete = SQLDelete.Replace("--BEGIN --[STIFFENER GROUP DELETE BEGIN]", "BEGIN --[STIFFENER GROUP DELETE BEGIN]")
'            SQLDelete = SQLDelete.Replace("--END --[STIFFENER GROUP DELETE END]", "END --[STIFFENER GROUP DELETE END]")
'            For Each row As StiffenerGroup In StiffenerGroups
'                SQLDelete = SQLDelete.Replace("--[STIFFENER GROUP INSERT]", row.SQLDelete)
'            Next
'        End If

'        Return SQLDelete

'    End Function

'#End Region

'#Region "Define"
'    Private _local_id As Integer?
'    Private _local_connection_id As Integer?
'    Private _ID As Integer?
'    Private _connection_id As Integer? 'currently called plate_id in EDS
'    Private _plate_location As String
'    Private _plate_type As String
'    Private _plate_diameter As Double?
'    Private _plate_thickness As Double?
'    Private _plate_material As Integer?
'    Private _stiffener_configuration As Integer?
'    Private _stiffener_clear_space As Double?
'    Private _plate_check As Boolean?

'    Public Property CCIplateMaterials As New List(Of CCIplateMaterial)
'    Public Property PlateResults As New List(Of PlateResults)
'    Public Property StiffenerGroups As New List(Of StiffenerGroup)

'    <Category("Plate Details"), Description(""), DisplayName("Local Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me._local_id
'        End Get
'        Set
'            Me._local_id = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Local Connection Id")>
'    Public Property local_connection_id() As Integer?
'        Get
'            Return Me._local_connection_id
'        End Get
'        Set
'            Me._local_connection_id = Value
'        End Set
'    End Property

'    <Category("Plate Details"), Description(""), DisplayName("Id")>
'    Public Property ID() As Integer?
'        Get
'            Return Me._ID
'        End Get
'        Set
'            Me._ID = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Connection Id")>
'    Public Property connection_id() As Integer?
'        Get
'            Return Me._connection_id
'        End Get
'        Set
'            Me._connection_id = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate Location")>
'    Public Property plate_location() As String
'        Get
'            Return Me._plate_location
'        End Get
'        Set
'            Me._plate_location = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate Type")>
'    Public Property plate_type() As String
'        Get
'            Return Me._plate_type
'        End Get
'        Set
'            Me._plate_type = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate Diameter")>
'    Public Property plate_diameter() As Double?
'        Get
'            Return Me._plate_diameter
'        End Get
'        Set
'            Me._plate_diameter = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate Thickness")>
'    Public Property plate_thickness() As Double?
'        Get
'            Return Me._plate_thickness
'        End Get
'        Set
'            Me._plate_thickness = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate Material")>
'    Public Property plate_material() As Integer?
'        Get
'            Return Me._plate_material
'        End Get
'        Set
'            Me._plate_material = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Stiffener Configuration")>
'    Public Property stiffener_configuration() As Integer?
'        Get
'            Return Me._stiffener_configuration
'        End Get
'        Set
'            Me._stiffener_configuration = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Stiffener Clear Space")>
'    Public Property stiffener_clear_space() As Double?
'        Get
'            Return Me._stiffener_clear_space
'        End Get
'        Set
'            Me._stiffener_clear_space = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate Check")>
'    Public Property plate_check() As Boolean?
'        Get
'            Return Me._plate_check
'        End Get
'        Set
'            Me._plate_check = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal pdrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = pdrow
'        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_plate_id")))
'        Me.ID = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_id = DBtoNullableInt(dr.Item("local_plate_id"))
'            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
'        End If
'        'Me.connection_id = If(EDStruefalse, DBtoNullableInt(dr.Item("plate_id")), DBtoNullableInt(dr.Item("local_connection_id"))) 'ME.plate_id '-pulls in null when Excel is referenced. 
'        If EDStruefalse = True Then 'only pull in when referencing EDS
'            Me.connection_id = DBtoNullableInt(dr.Item("plate_id"))
'        End If
'        Me.plate_location = DBtoStr(dr.Item("plate_location"))
'        Me.plate_type = DBtoStr(dr.Item("plate_type"))
'        Me.plate_diameter = DBtoNullableDbl(dr.Item("plate_diameter"))
'        Me.plate_thickness = DBtoNullableDbl(dr.Item("plate_thickness"))
'        Me.plate_material = DBtoNullableInt(dr.Item("plate_material"))
'        Me.stiffener_configuration = If(DBtoStr(dr.Item("stiffener_configuration")) = "Custom", 4, DBtoNullableInt(dr.Item("stiffener_configuration"))) 'Stiffener configuration is 0 through 3 plus 'custom'. Custom will report as option 4. 
'        Me.stiffener_clear_space = DBtoNullableDbl(dr.Item("stiffener_clear_space"))
'        Me.plate_check = If(EDStruefalse, DBtoNullableBool(dr.Item("plate_check")), If(DBtoStr(dr.Item("plate_check")) = "Yes", True, If(DBtoStr(dr.Item("plate_check")) = "No", False, DBtoNullableBool(dr.Item("plate_check"))))) 'Listed as a string and need to convert to Boolean

'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_location.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_type.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_diameter.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_thickness.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_material.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_configuration.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_clear_space.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_check.ToString.FormatDBValue)

'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_id")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_location")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_type")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_diameter")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_thickness")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_material")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_configuration")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_clear_space")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_check")

'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_id = " & "@SubLevel1ID")
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_location = " & Me.plate_location.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_type = " & Me.plate_type.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_diameter = " & Me.plate_diameter.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_thickness = " & Me.plate_thickness.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_material = " & Me.plate_material.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_material = " & "@SubLevel3ID")
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_configuration = " & Me.stiffener_configuration.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_clear_space = " & Me.stiffener_clear_space.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_check = " & Me.plate_check.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'plate_material references the local id when coming from Excel. Need to convert to EDS ID when performing Equals function
'        Dim material As Integer?
'        For Each row As CCIplateMaterial In CCIplateMaterials
'            If Me.plate_material = row.local_id And row.ID > 0 Then
'                material = row.ID
'                Exit For
'            End If
'        Next


'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As PlateDetail = TryCast(other, PlateDetail)
'        If otherToCompare Is Nothing Then Return False

'        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
'        'Equals = If(Me.plate_id.CheckChange(otherToCompare.plate_id, changes, categoryName, "Plate Id"), Equals, False)
'        Equals = If(Me.plate_location.CheckChange(otherToCompare.plate_location, changes, categoryName, "Plate Location"), Equals, False)
'        Equals = If(Me.plate_type.CheckChange(otherToCompare.plate_type, changes, categoryName, "Plate Type"), Equals, False)
'        Equals = If(Me.plate_diameter.CheckChange(otherToCompare.plate_diameter, changes, categoryName, "Plate Diameter"), Equals, False)
'        Equals = If(Me.plate_thickness.CheckChange(otherToCompare.plate_thickness, changes, categoryName, "Plate Thickness"), Equals, False)
'        'Equals = If(Me.plate_material.CheckChange(otherToCompare.plate_material, changes, categoryName, "Plate Material"), Equals, False)
'        Equals = If(material.CheckChange(otherToCompare.plate_material, changes, categoryName, "Plate Material"), Equals, False)
'        Equals = If(Me.stiffener_configuration.CheckChange(otherToCompare.stiffener_configuration, changes, categoryName, "Stiffener Configuration"), Equals, False)
'        Equals = If(Me.stiffener_clear_space.CheckChange(otherToCompare.stiffener_clear_space, changes, categoryName, "Stiffener Clear Space"), Equals, False)
'        Equals = If(Me.plate_check.CheckChange(otherToCompare.plate_check, changes, categoryName, "Plate Check"), Equals, False)

'        'Materials
'        If Me.CCIplateMaterials.Count > 0 Then
'            Equals = If(Me.CCIplateMaterials.CheckChange(otherToCompare.CCIplateMaterials, changes, categoryName, "CCIplate Materials"), Equals, False)
'        End If

'        'Stiffener Groups
'        If Me.StiffenerGroups.Count > 0 Then
'            Equals = If(Me.StiffenerGroups.CheckChange(otherToCompare.StiffenerGroups, changes, categoryName, "Stiffener Groups"), Equals, False)
'        End If

'    End Function
'#End Region

'End Class
'Partial Public Class BoltGroup
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Bolt Groups"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.bolts"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Group (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Group_INSERT
'        SQLInsert = SQLInsert.Replace("[BOLT GROUP VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[BOLT GROUP FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Bolt Detail
'        If Me.BoltDetails.Count > 0 Then
'            SQLInsert = SQLInsert.Replace("--BEGIN --[BOLT DETAIL INSERT BEGIN]", "BEGIN --[BOLT DETAIL INSERT BEGIN]")
'            SQLInsert = SQLInsert.Replace("--END --[BOLT DETAIL INSERT END]", "END --[BOLT DETAIL INSERT END]")
'            For Each row As BoltDetail In BoltDetails
'                SQLInsert = SQLInsert.Replace("--[BOLT DETAIL INSERT]", row.SQLInsert)
'            Next
'        End If

'        'Results
'        For Each row As BoltResults In BoltResults
'            SQLInsert = SQLInsert.Replace("--BEGIN --[BOLT RESULTS INSERT BEGIN]", "BEGIN --[BOLT RESULTS INSERT BEGIN]")
'            SQLInsert = SQLInsert.Replace("--END --[BOLT RESULTS INSERT END]", "END --[BOLT RESULTS INSERT END]")
'            SQLInsert = SQLInsert.Replace("--[BOLT RESULTS INSERT]", row.SQLInsert)
'        Next

'        Return SQLInsert

'    End Function

'    Public Overrides Function SQLUpdate() As String

'        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Group (UPDATE).sql")
'        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Group_UPDATE
'        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
'        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Bolt Detail
'        If Me.BoltDetails.Count > 0 Then
'            SQLUpdate = SQLUpdate.Replace("--BEGIN --[BOLT DETAIL UPDATE BEGIN]", "BEGIN --[BOLT DETAIL UPDATE BEGIN]")
'            SQLUpdate = SQLUpdate.Replace("--END --[BOLT DETAIL UPDATE END]", "END --[BOLT DETAIL UPDATE END]")
'            For Each row As BoltDetail In BoltDetails
'                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
'                    If IsSomething(row.bolt_location) Or IsSomething(row.bolt_diameter) Or IsSomething(row.bolt_material) Or IsSomething(row.bolt_circle) Or IsSomething(row.eta_factor) Or IsSomething(row.lar) Or IsSomethingString(row.bolt_thread_type) Or IsSomething(row.area_override) Or IsSomething(row.tension_only) Then
'                        SQLUpdate = SQLUpdate.Replace("--[BOLT DETAIL INSERT]", row.SQLUpdate)
'                    Else
'                        SQLUpdate = SQLUpdate.Replace("--[BOLT DETAIL INSERT]", row.SQLDelete)
'                    End If
'                Else
'                    SQLUpdate = SQLUpdate.Replace("--[BOLT DETAIL INSERT]", row.SQLInsert)
'                End If
'            Next
'        End If

'        For Each row As BoltResults In BoltResults
'            SQLUpdate = SQLUpdate.Replace("--BEGIN --[BOLT RESULTS INSERT BEGIN]", "BEGIN --[BOLT RESULTS INSERT BEGIN]")
'            SQLUpdate = SQLUpdate.Replace("--END --[BOLT RESULTS INSERT END]", "END --[BOLT RESULTS INSERT END]")
'            SQLUpdate = SQLUpdate.Replace("--[BOLT RESULTS INSERT]", row.SQLInsert)
'        Next

'        Return SQLUpdate

'    End Function

'    Public Overrides Function SQLDelete() As String

'        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Group (DELETE).sql")
'        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Group_DELETE
'        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Bolt Details
'        If Me.BoltDetails.Count > 0 Then
'            SQLDelete = SQLDelete.Replace("--BEGIN --[BOLT DETAIL DELETE BEGIN]", "BEGIN --[BOLT DETAIL DELETE BEGIN]")
'            SQLDelete = SQLDelete.Replace("--END --[BOLT DETAIL DELETE END]", "END --[BOLT DETAIL DELETE END]")
'            For Each row As BoltDetail In BoltDetails
'                SQLDelete = SQLDelete.Replace("--[BOLT DETAIL INSERT]", row.SQLDelete)
'            Next
'        End If

'        Return SQLDelete

'    End Function

'#End Region

'#Region "Define"
'    Private _local_id As Integer?
'    Private _local_connection_id As Integer?
'    Private _ID As Integer?
'    Private _connection_id As Integer?
'    Private _resist_axial As Boolean?
'    Private _resist_shear As Boolean?
'    Private _plate_bending As Boolean?
'    Private _grout_considered As Boolean?
'    Private _apply_barb_elevation As Boolean?
'    Private _bolt_name As String


'    Public Property BoltDetails As New List(Of BoltDetail)
'    Public Property BoltResults As New List(Of BoltResults)

'    <Category("Bolt Groups"), Description(""), DisplayName("Local Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me._local_id
'        End Get
'        Set
'            Me._local_id = Value
'        End Set
'    End Property
'    <Category("Bolt Groups"), Description(""), DisplayName("Local Connection Id")>
'    Public Property local_connection_id() As Integer?
'        Get
'            Return Me._local_connection_id
'        End Get
'        Set
'            Me._local_connection_id = Value
'        End Set
'    End Property

'    <Category("Bolt Groups"), Description(""), DisplayName("Id")>
'    Public Property ID() As Integer?
'        Get
'            Return Me._ID
'        End Get
'        Set
'            Me._ID = Value
'        End Set
'    End Property
'    <Category("Bolt Groups"), Description(""), DisplayName("Connection Id")>
'    Public Property connection_id() As Integer?
'        Get
'            Return Me._connection_id
'        End Get
'        Set
'            Me._connection_id = Value
'        End Set
'    End Property
'    <Category("Bolt Groups"), Description(""), DisplayName("Resist Axial")>
'    Public Property resist_axial() As Boolean?
'        Get
'            Return Me._resist_axial
'        End Get
'        Set
'            Me._resist_axial = Value
'        End Set
'    End Property
'    <Category("Bolt Groups"), Description(""), DisplayName("Resist Shear")>
'    Public Property resist_shear() As Boolean?
'        Get
'            Return Me._resist_shear
'        End Get
'        Set
'            Me._resist_shear = Value
'        End Set
'    End Property
'    <Category("Bolt Groups"), Description(""), DisplayName("Plate Bending")>
'    Public Property plate_bending() As Boolean?
'        Get
'            Return Me._plate_bending
'        End Get
'        Set
'            Me._plate_bending = Value
'        End Set
'    End Property
'    <Category("Bolt Groups"), Description(""), DisplayName("Grout Considered")>
'    Public Property grout_considered() As Boolean?
'        Get
'            Return Me._grout_considered
'        End Get
'        Set
'            Me._grout_considered = Value
'        End Set
'    End Property
'    <Category("Bolt Groups"), Description(""), DisplayName("Apply Barb Elevation")>
'    Public Property apply_barb_elevation() As Boolean?
'        Get
'            Return Me._apply_barb_elevation
'        End Get
'        Set
'            Me._apply_barb_elevation = Value
'        End Set
'    End Property
'    <Category("Bolt Groups"), Description(""), DisplayName("Bolt Name")>
'    Public Property bolt_name() As String
'        Get
'            Return Me._bolt_name
'        End Get
'        Set
'            Me._bolt_name = Value
'        End Set
'    End Property


'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal bgrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = bgrow
'        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_plate_id")))
'        Me.ID = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_id = DBtoNullableInt(dr.Item("local_bolt_group_id"))
'            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
'        End If
'        'Me.connection_id = If(EDStruefalse, DBtoNullableInt(dr.Item("plate_id")), DBtoNullableInt(dr.Item("local_connection_id")))
'        If EDStruefalse = True Then 'only pull in when referencing EDS
'            Me.connection_id = DBtoNullableInt(dr.Item("plate_id"))
'        End If
'        Me.resist_axial = If(EDStruefalse, DBtoNullableBool(dr.Item("resist_axial")), If(DBtoStr(dr.Item("resist_axial")) = "Yes", True, If(DBtoStr(dr.Item("resist_axial")) = "No", False, DBtoNullableBool(dr.Item("resist_axial")))))
'        Me.resist_shear = If(EDStruefalse, DBtoNullableBool(dr.Item("resist_shear")), If(DBtoStr(dr.Item("resist_shear")) = "Yes", True, If(DBtoStr(dr.Item("resist_shear")) = "No", False, DBtoNullableBool(dr.Item("resist_shear")))))
'        Me.plate_bending = If(EDStruefalse, DBtoNullableBool(dr.Item("plate_bending")), If(DBtoStr(dr.Item("plate_bending")) = "Yes", True, If(DBtoStr(dr.Item("plate_bending")) = "No", False, DBtoNullableBool(dr.Item("plate_bending")))))
'        Me.grout_considered = If(EDStruefalse, DBtoNullableBool(dr.Item("grout_considered")), If(DBtoStr(dr.Item("grout_considered")) = "Yes", True, If(DBtoStr(dr.Item("grout_considered")) = "No", False, DBtoNullableBool(dr.Item("grout_considered")))))
'        Me.apply_barb_elevation = If(EDStruefalse, DBtoNullableBool(dr.Item("apply_barb_elevation")), If(DBtoStr(dr.Item("apply_barb_elevation")) = "Yes", True, If(DBtoStr(dr.Item("apply_barb_elevation")) = "No", False, DBtoNullableBool(dr.Item("apply_barb_elevation")))))
'        Me.bolt_name = If(EDStruefalse, DBtoStr(dr.Item("bolt_name")), DBtoStr(dr.Item("group_name")))

'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_id.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.resist_axial.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.resist_shear.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_bending.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.grout_considered.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.apply_barb_elevation.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_name.ToString.FormatDBValue)

'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_id")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("resist_axial")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("resist_shear")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_bending")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("grout_considered")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("apply_barb_elevation")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_name")

'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_id = " & "@SubLevel1ID")
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("resist_axial = " & Me.resist_axial.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("resist_shear = " & Me.resist_shear.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_bending = " & Me.plate_bending.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("grout_considered = " & Me.grout_considered.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("apply_barb_elevation = " & Me.apply_barb_elevation.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_name = " & Me.bolt_name.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As BoltGroup = TryCast(other, BoltGroup)
'        If otherToCompare Is Nothing Then Return False

'        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
'        'Equals = If(Me.plate_id.CheckChange(otherToCompare.plate_id, changes, categoryName, "Plate Id"), Equals, False)
'        Equals = If(Me.resist_axial.CheckChange(otherToCompare.resist_axial, changes, categoryName, "Resist Axial"), Equals, False)
'        Equals = If(Me.resist_shear.CheckChange(otherToCompare.resist_shear, changes, categoryName, "Resist Shear"), Equals, False)
'        Equals = If(Me.plate_bending.CheckChange(otherToCompare.plate_bending, changes, categoryName, "Plate Bending"), Equals, False)
'        Equals = If(Me.grout_considered.CheckChange(otherToCompare.grout_considered, changes, categoryName, "Grout Considered"), Equals, False)
'        Equals = If(Me.apply_barb_elevation.CheckChange(otherToCompare.apply_barb_elevation, changes, categoryName, "Apply Barb Elevation"), Equals, False)
'        Equals = If(Me.bolt_name.CheckChange(otherToCompare.bolt_name, changes, categoryName, "Bolt Name"), Equals, False)

'        'Bolt Details
'        If Me.BoltDetails.Count > 0 Then
'            Equals = If(Me.BoltDetails.CheckChange(otherToCompare.BoltDetails, changes, categoryName, "Bolt Details"), Equals, False)
'        End If

'    End Function
'#End Region

'End Class
'Partial Public Class BoltDetail
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Bolt Details"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.bolt_details"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Detail (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Detail_INSERT
'        SQLInsert = SQLInsert.Replace("[BOLT DETAIL VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[BOLT DETAIL FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Plate Material
'        For Each row As CCIplateMaterial In CCIplateMaterials
'            SQLInsert = SQLInsert.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert)
'        Next

'        Return SQLInsert

'    End Function

'    Public Overrides Function SQLUpdate() As String

'        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Detail (UPDATE).sql")
'        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Detail_UPDATE
'        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
'        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Plate Material
'        For Each row As CCIplateMaterial In CCIplateMaterials
'            SQLUpdate = SQLUpdate.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert) 'Can only insert materials, no deleting or updating since database is referenced by all BUs. 
'        Next

'        Return SQLUpdate

'    End Function

'    Public Overrides Function SQLDelete() As String

'        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Detail (DELETE).sql")
'        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Detail_DELETE
'        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLDelete

'    End Function

'#End Region

'#Region "Define"
'    'Private _local_id As Integer?
'    Private _local_connection_id As Integer?
'    Private _local_group_id As Integer?
'    Private _ID As Integer?
'    Private _bolt_group_id As Integer?
'    Private _bolt_location As Double?
'    Private _bolt_diameter As Double?
'    Private _bolt_material As Integer?
'    Private _bolt_circle As Double?
'    Private _eta_factor As Double?
'    Private _lar As Double?
'    Private _bolt_thread_type As String
'    Private _area_override As Double?
'    Private _tension_only As Boolean?

'    Public Property CCIplateMaterials As New List(Of CCIplateMaterial)

'    '<Category("Bolt Details"), Description(""), DisplayName("Local Id")>
'    'Public Property local_id() As Integer?
'    '    Get
'    '        Return Me._local_id
'    '    End Get
'    '    Set
'    '        Me._local_id = Value
'    '    End Set
'    'End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Local Connection Id")>
'    Public Property local_connection_id() As Integer?
'        Get
'            Return Me._local_connection_id
'        End Get
'        Set
'            Me._local_connection_id = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Local Group Id")>
'    Public Property local_group_id() As Integer?
'        Get
'            Return Me._local_group_id
'        End Get
'        Set
'            Me._local_group_id = Value
'        End Set
'    End Property

'    <Category("Bolt Details"), Description(""), DisplayName("Id")>
'    Public Property ID() As Integer?
'        Get
'            Return Me._ID
'        End Get
'        Set
'            Me._ID = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Bolt Group Id")>
'    Public Property bolt_group_id() As Integer?
'        Get
'            Return Me._bolt_group_id
'        End Get
'        Set
'            Me._bolt_group_id = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Bolt Location")>
'    Public Property bolt_location() As Double?
'        Get
'            Return Me._bolt_location
'        End Get
'        Set
'            Me._bolt_location = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Bolt Diameter")>
'    Public Property bolt_diameter() As Double?
'        Get
'            Return Me._bolt_diameter
'        End Get
'        Set
'            Me._bolt_diameter = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Bolt Material")>
'    Public Property bolt_material() As Integer?
'        Get
'            Return Me._bolt_material
'        End Get
'        Set
'            Me._bolt_material = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Bolt Circle")>
'    Public Property bolt_circle() As Double?
'        Get
'            Return Me._bolt_circle
'        End Get
'        Set
'            Me._bolt_circle = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Eta Factor")>
'    Public Property eta_factor() As Double?
'        Get
'            Return Me._eta_factor
'        End Get
'        Set
'            Me._eta_factor = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Lar")>
'    Public Property lar() As Double?
'        Get
'            Return Me._lar
'        End Get
'        Set
'            Me._lar = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Bolt Thread Type")>
'    Public Property bolt_thread_type() As String
'        Get
'            Return Me._bolt_thread_type
'        End Get
'        Set
'            Me._bolt_thread_type = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Area Override")>
'    Public Property area_override() As Double?
'        Get
'            Return Me._area_override
'        End Get
'        Set
'            Me._area_override = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Tension Only")>
'    Public Property tension_only() As Boolean?
'        Get
'            Return Me._tension_only
'        End Get
'        Set
'            Me._tension_only = Value
'        End Set
'    End Property


'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal bdrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = bdrow
'        Me.ID = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            'Me.local_id = DBtoNullableInt(dr.Item("local_connection_id")) 'nothing references bolt details so deactivating
'            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
'            Me.local_group_id = DBtoNullableInt(dr.Item("local_group_id"))
'        End If
'        'Me.bolt_id = If(EDStruefalse, DBtoNullableInt(dr.Item("bolt_id")), DBtoNullableInt(dr.Item("local_group_id")))
'        Me.bolt_group_id = If(EDStruefalse, DBtoNullableInt(dr.Item("bolt_id")), DBtoNullableInt(dr.Item("group_id")))
'        Me.bolt_location = DBtoNullableDbl(dr.Item("bolt_location"))
'        Me.bolt_diameter = DBtoNullableDbl(dr.Item("bolt_diameter"))
'        Me.bolt_material = DBtoNullableInt(dr.Item("bolt_material"))
'        Me.bolt_circle = DBtoNullableDbl(dr.Item("bolt_circle"))
'        Me.eta_factor = DBtoNullableDbl(dr.Item("eta_factor"))
'        Me.lar = DBtoNullableDbl(dr.Item("lar"))
'        Me.bolt_thread_type = DBtoStr(dr.Item("bolt_thread_type"))
'        Me.area_override = DBtoNullableDbl(dr.Item("area_override"))
'        'When data is coming from Excel, blank data will report nothing. 
'        Me.tension_only = If(EDStruefalse, DBtoNullableBool(dr.Item("tension_only")), If(DBtoStr(dr.Item("tension_only")) = "Yes", True, If(DBtoStr(dr.Item("tension_only")) = "No", False, DBtoNullableBool(dr.Item("tension_only")))))


'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_id.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_location.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_diameter.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_material.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_circle.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.eta_factor.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lar.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_thread_type.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.area_override.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tension_only.ToString.FormatDBValue)

'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_id")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_location")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_diameter")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_material")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_circle")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("eta_factor")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("lar")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_thread_type")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("area_override")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("tension_only")

'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_id = " & "@SubLevel2ID")
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_location = " & Me.bolt_location.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_diameter = " & Me.bolt_diameter.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_material = " & Me.bolt_material.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_material = " & "@SubLevel3ID")
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_circle = " & Me.bolt_circle.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("eta_factor = " & Me.eta_factor.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lar = " & Me.lar.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_thread_type = " & Me.bolt_thread_type.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("area_override = " & Me.area_override.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tension_only = " & Me.tension_only.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As BoltDetail = TryCast(other, BoltDetail)
'        If otherToCompare Is Nothing Then Return False

'        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
'        Equals = If(Me.bolt_group_id.CheckChange(otherToCompare.bolt_group_id, changes, categoryName, "Bolt Group Id"), Equals, False)
'        Equals = If(Me.bolt_location.CheckChange(otherToCompare.bolt_location, changes, categoryName, "Bolt Location"), Equals, False)
'        Equals = If(Me.bolt_diameter.CheckChange(otherToCompare.bolt_diameter, changes, categoryName, "Bolt Diameter"), Equals, False)
'        Equals = If(Me.bolt_material.CheckChange(otherToCompare.bolt_material, changes, categoryName, "Bolt Material"), Equals, False)
'        Equals = If(Me.bolt_circle.CheckChange(otherToCompare.bolt_circle, changes, categoryName, "Bolt Circle"), Equals, False)
'        Equals = If(Me.eta_factor.CheckChange(otherToCompare.eta_factor, changes, categoryName, "Eta Factor"), Equals, False)
'        Equals = If(Me.lar.CheckChange(otherToCompare.lar, changes, categoryName, "Lar"), Equals, False)
'        Equals = If(Me.bolt_thread_type.CheckChange(otherToCompare.bolt_thread_type, changes, categoryName, "Bolt Thread Type"), Equals, False)
'        Equals = If(Me.area_override.CheckChange(otherToCompare.area_override, changes, categoryName, "Area Override"), Equals, False)
'        Equals = If(Me.tension_only.CheckChange(otherToCompare.tension_only, changes, categoryName, "Tension Only"), Equals, False)

'        'Materials
'        If Me.CCIplateMaterials.Count > 0 Then
'            Equals = If(Me.CCIplateMaterials.CheckChange(otherToCompare.CCIplateMaterials, changes, categoryName, "CCIplate Materials"), Equals, False)
'        End If

'    End Function
'#End Region

'End Class

'Partial Public Class CCIplateMaterial
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "CCIplate Materials"
'    Public Overrides ReadOnly Property EDSTableName As String = "gen.connection_material_properties"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate Material (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Material_INSERT
'        SQLInsert = SQLInsert.Replace("[MATERIAL PROPERTY ID]", Me.ID.ToString.FormatDBValue)
'        SQLInsert = SQLInsert.Replace("[SELECT]", Me.SQLUpdateFieldsandValues)
'        SQLInsert = SQLInsert.Replace("[CCIPLATE MATERIAL VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[CCIPLATE MATERIAL FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLInsert

'    End Function

'#End Region

'#Region "Define"
'    Private _ID As Integer?
'    Private _local_id As Integer? 'removed
'    Private _name As String
'    Private _fy_0 As Double?
'    Private _fy_1_125 As Double?
'    Private _fy_1_625 As Double?
'    Private _fy_2_625 As Double?
'    Private _fy_4_125 As Double?
'    Private _fu_0 As Double?
'    Private _fu_1_125 As Double?
'    Private _fu_1_625 As Double?
'    Private _fu_2_625 As Double?
'    Private _fu_4_125 As Double?
'    Private _default_material As Boolean?
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Id")>
'    Public Property ID() As Integer?
'        Get
'            Return Me._ID
'        End Get
'        Set
'            Me._ID = Value
'        End Set
'    End Property
'    <Category("Connection Material Properties"), Description(""), DisplayName("Local Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me._local_id
'        End Get
'        Set
'            Me._local_id = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Name")>
'    Public Property name() As String
'        Get
'            Return Me._name
'        End Get
'        Set
'            Me._name = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 0")>
'    Public Property fy_0() As Double?
'        Get
'            Return Me._fy_0
'        End Get
'        Set
'            Me._fy_0 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 1 125")>
'    Public Property fy_1_125() As Double?
'        Get
'            Return Me._fy_1_125
'        End Get
'        Set
'            Me._fy_1_125 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 1 625")>
'    Public Property fy_1_625() As Double?
'        Get
'            Return Me._fy_1_625
'        End Get
'        Set
'            Me._fy_1_625 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 2 625")>
'    Public Property fy_2_625() As Double?
'        Get
'            Return Me._fy_2_625
'        End Get
'        Set
'            Me._fy_2_625 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 4 125")>
'    Public Property fy_4_125() As Double?
'        Get
'            Return Me._fy_4_125
'        End Get
'        Set
'            Me._fy_4_125 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 0")>
'    Public Property fu_0() As Double?
'        Get
'            Return Me._fu_0
'        End Get
'        Set
'            Me._fu_0 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 1 125")>
'    Public Property fu_1_125() As Double?
'        Get
'            Return Me._fu_1_125
'        End Get
'        Set
'            Me._fu_1_125 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 1 625")>
'    Public Property fu_1_625() As Double?
'        Get
'            Return Me._fu_1_625
'        End Get
'        Set
'            Me._fu_1_625 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 2 625")>
'    Public Property fu_2_625() As Double?
'        Get
'            Return Me._fu_2_625
'        End Get
'        Set
'            Me._fu_2_625 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 4 125")>
'    Public Property fu_4_125() As Double?
'        Get
'            Return Me._fu_4_125
'        End Get
'        Set
'            Me._fu_4_125 = Value
'        End Set
'    End Property
'    <Category("CCIplate Material Properties"), Description(""), DisplayName("Default Material")>
'    Public Property default_material() As Boolean?
'        Get
'            Return Me._default_material
'        End Get
'        Set
'            Me._default_material = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal mrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal mrow As DataRow, ByRef strDS As DataSet)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = mrow
'        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_id")))
'        Me.ID = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_id = DBtoNullableInt(dr.Item("local_material_id"))
'        End If
'        Me.name = DBtoStr(dr.Item("name"))
'        Me.fy_0 = DBtoNullableDbl(dr.Item("fy_0"))
'        Me.fy_1_125 = DBtoNullableDbl(dr.Item("fy_1_125"))
'        Me.fy_1_625 = DBtoNullableDbl(dr.Item("fy_1_625"))
'        Me.fy_2_625 = DBtoNullableDbl(dr.Item("fy_2_625"))
'        Me.fy_4_125 = DBtoNullableDbl(dr.Item("fy_4_125"))
'        Me.fu_0 = DBtoNullableDbl(dr.Item("fu_0"))
'        Me.fu_1_125 = DBtoNullableDbl(dr.Item("fu_1_125"))
'        Me.fu_1_625 = DBtoNullableDbl(dr.Item("fu_1_625"))
'        Me.fu_2_625 = DBtoNullableDbl(dr.Item("fu_2_625"))
'        Me.fu_4_125 = DBtoNullableDbl(dr.Item("fu_4_125"))
'        Me.default_material = If(EDStruefalse, DBtoNullableBool(dr.Item("default_material")), False)


'    End Sub

'    Public Sub New(ByVal ID As Integer?)
'        'This is used to store a temp list of new materials to add to the Excel tool
'        Me.ID = ID
'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_id.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.name.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_0.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_1_125.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_1_625.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_2_625.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_4_125.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_0.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_1_125.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_1_625.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_2_625.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_4_125.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.default_material.ToString.FormatDBValue)

'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("local_id")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("name")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_0")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_1_125")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_1_625")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_2_625")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_4_125")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_0")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_1_125")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_1_625")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_2_625")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_4_125")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("default_material")

'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("local_id = " & Me.local_id.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("name = " & Me.name.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_0), "fy_0 = " & Me.fy_0.ToString.FormatDBValue, "fy_0 IS NULL "))
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_1_125), "fy_1_125 = " & Me.fy_1_125.ToString.FormatDBValue, "fy_1_125 IS NULL "))
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_1_625), "fy_1_625 = " & Me.fy_1_625.ToString.FormatDBValue, "fy_1_625 IS NULL "))
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_2_625), "fy_2_625 = " & Me.fy_2_625.ToString.FormatDBValue, "fy_2_625 IS NULL "))
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_4_125), "fy_4_125 = " & Me.fy_4_125.ToString.FormatDBValue, "fy_4_125 IS NULL "))
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_0), "fu_0 = " & Me.fu_0.ToString.FormatDBValue, "fu_0 IS NULL "))
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_1_125), "fu_1_125 = " & Me.fu_1_125.ToString.FormatDBValue, "fu_1_125 IS NULL "))
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_1_625), "fu_1_625 = " & Me.fu_1_625.ToString.FormatDBValue, "fu_1_625 IS NULL "))
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_2_625), "fu_2_625 = " & Me.fu_2_625.ToString.FormatDBValue, "fu_2_625 IS NULL "))
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_4_125), "fu_4_125 = " & Me.fu_4_125.ToString.FormatDBValue, "fu_4_125 IS NULL "))

'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("default_material = " & Me.default_material.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As CCIplateMaterial = TryCast(other, CCIplateMaterial)
'        If otherToCompare Is Nothing Then Return False

'        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
'        'Equals = If(Me.local_id.CheckChange(otherToCompare.local_id, changes, categoryName, "Local Id"), Equals, False)
'        Equals = If(Me.name.CheckChange(otherToCompare.name, changes, categoryName, "Name"), Equals, False)
'        Equals = If(Me.fy_0.CheckChange(otherToCompare.fy_0, changes, categoryName, "Fy 0"), Equals, False)
'        Equals = If(Me.fy_1_125.CheckChange(otherToCompare.fy_1_125, changes, categoryName, "Fy 1 125"), Equals, False)
'        Equals = If(Me.fy_1_625.CheckChange(otherToCompare.fy_1_625, changes, categoryName, "Fy 1 625"), Equals, False)
'        Equals = If(Me.fy_2_625.CheckChange(otherToCompare.fy_2_625, changes, categoryName, "Fy 2 625"), Equals, False)
'        Equals = If(Me.fy_4_125.CheckChange(otherToCompare.fy_4_125, changes, categoryName, "Fy 4 125"), Equals, False)
'        Equals = If(Me.fu_0.CheckChange(otherToCompare.fu_0, changes, categoryName, "Fu 0"), Equals, False)
'        Equals = If(Me.fu_1_125.CheckChange(otherToCompare.fu_1_125, changes, categoryName, "Fu 1 125"), Equals, False)
'        Equals = If(Me.fu_1_625.CheckChange(otherToCompare.fu_1_625, changes, categoryName, "Fu 1 625"), Equals, False)
'        Equals = If(Me.fu_2_625.CheckChange(otherToCompare.fu_2_625, changes, categoryName, "Fu 2 625"), Equals, False)
'        Equals = If(Me.fu_4_125.CheckChange(otherToCompare.fu_4_125, changes, categoryName, "Fu 4 125"), Equals, False)
'        'Equals = If(Me.default_material.CheckChange(otherToCompare.default_material, changes, categoryName, "Default Material"), Equals, False)


'    End Function
'#End Region

'End Class

'Partial Public Class PlateResults
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Plate Results"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.plate_results"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Plate Result (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Plate_Result_INSERT
'        SQLInsert = SQLInsert.Replace("[PLATE RESULT VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[PLATE RESULT FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLInsert

'    End Function

'#End Region

'#Region "Define"
'    Private _plate_details_id As Integer?
'    Private _local_plate_id As Integer?
'    'Private _work_order_seq_num As Double? 'not provided in Excel
'    Private _rating As Double?
'    Private _result_lkup As String
'    'Private _modified_person_id As Integer? 'not provided in Excel
'    'Private _process_stage As String 'not provided in Excel
'    'Private _modified_date As DateTime? 'not provided in Excel

'    <Category("Plate Results"), Description(""), DisplayName("Plate Details Id")>
'    Public Property plate_details_id() As Integer?
'        Get
'            Return Me._plate_details_id
'        End Get
'        Set
'            Me._plate_details_id = Value
'        End Set
'    End Property
'    <Category("Plate Results"), Description(""), DisplayName("Local Plate Id")>
'    Public Property local_plate_id() As Integer?
'        Get
'            Return Me._local_plate_id
'        End Get
'        Set
'            Me._local_plate_id = Value
'        End Set
'    End Property
'    '<Category("Plate Results"), Description(""), DisplayName("Work Order Seq Num")>
'    'Public Property work_order_seq_num() As Double?
'    '    Get
'    '        Return Me._work_order_seq_num
'    '    End Get
'    '    Set
'    '        Me._work_order_seq_num = Value
'    '    End Set
'    'End Property
'    <Category("Plate Results"), Description(""), DisplayName("Rating")>
'    Public Property rating() As Double?
'        Get
'            Return Me._rating
'        End Get
'        Set
'            Me._rating = Value
'        End Set
'    End Property
'    <Category("Plate Results"), Description(""), DisplayName("Result Lkup")>
'    Public Property result_lkup() As String
'        Get
'            Return Me._result_lkup
'        End Get
'        Set
'            Me._result_lkup = Value
'        End Set
'    End Property
'    '<Category("Plate Results"), Description(""), DisplayName("Modified Person Id")>
'    'Public Property modified_person_id() As Integer?
'    '    Get
'    '        Return Me._modified_person_id
'    '    End Get
'    '    Set
'    '        Me._modified_person_id = Value
'    '    End Set
'    'End Property
'    '<Category("Plate Results"), Description(""), DisplayName("Process Stage")>
'    'Public Property process_stage() As String
'    '    Get
'    '        Return Me._process_stage
'    '    End Get
'    '    Set
'    '        Me._process_stage = Value
'    '    End Set
'    'End Property
'    '<Category("Plate Results"), Description(""), DisplayName("Modified Date")>
'    'Public Property modified_date() As DateTime?
'    '    Get
'    '        Return Me._modified_date
'    '    End Get
'    '    Set
'    '        Me._modified_date = Value
'    '    End Set
'    'End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal prrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = prrow

'        Me.plate_details_id = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_plate_id = DBtoNullableInt(dr.Item("local_plate_id"))
'        End If
'        'Me.work_order_seq_num = DBtoNullableDbl(dr.Item("work_order_seq_num"))
'        Me.rating = DBtoNullableDbl(dr.Item("rating"))
'        Me.result_lkup = DBtoStr(dr.Item("result_lkup"))
'        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
'        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
'        'Me.modified_date = DBtoStr(dr.Item("modified_date"))

'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""

'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)


'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""

'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_details_id")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")


'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_details_id = " & Me.plate_details_id.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As PlateResults = TryCast(other, PlateResults)
'        If otherToCompare Is Nothing Then Return False


'    End Function
'#End Region

'End Class

'Partial Public Class BoltResults
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Bolt Results"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.bolt_results"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Result (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Result_INSERT
'        SQLInsert = SQLInsert.Replace("[BOLT RESULT VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[BOLT RESULT FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLInsert

'    End Function

'#End Region

'#Region "Define"
'    Private _bolt_id As Integer?
'    Private _local_connection_id As Integer?
'    Private _local_bolt_group_id As Integer?
'    'Private _work_order_seq_num As Double? 'not provided in Excel
'    Private _rating As Double?
'    Private _result_lkup As String
'    'Private _modified_person_id As Integer? 'not provided in Excel
'    'Private _process_stage As String 'not provided in Excel
'    'Private _modified_date As DateTime? 'not provided in Excel

'    <Category("Bolt Results"), Description(""), DisplayName("Bolt Id")>
'    Public Property bolt_id() As Integer?
'        Get
'            Return Me._bolt_id
'        End Get
'        Set
'            Me._bolt_id = Value
'        End Set
'    End Property
'    <Category("Bolt Results"), Description(""), DisplayName("Local Connection Id")>
'    Public Property local_connection_id() As Integer?
'        Get
'            Return Me._local_connection_id
'        End Get
'        Set
'            Me._local_connection_id = Value
'        End Set
'    End Property
'    <Category("Bolt Results"), Description(""), DisplayName("Local Bolt Group Id")>
'    Public Property local_bolt_group_id() As Integer?
'        Get
'            Return Me._local_bolt_group_id
'        End Get
'        Set
'            Me._local_bolt_group_id = Value
'        End Set
'    End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Work Order Seq Num")>
'    'Public Property work_order_seq_num() As Double?
'    '    Get
'    '        Return Me._work_order_seq_num
'    '    End Get
'    '    Set
'    '        Me._work_order_seq_num = Value
'    '    End Set
'    'End Property
'    <Category("Bolt Results"), Description(""), DisplayName("Rating")>
'    Public Property rating() As Double?
'        Get
'            Return Me._rating
'        End Get
'        Set
'            Me._rating = Value
'        End Set
'    End Property
'    <Category("Bolt Results"), Description(""), DisplayName("Result Lkup")>
'    Public Property result_lkup() As String
'        Get
'            Return Me._result_lkup
'        End Get
'        Set
'            Me._result_lkup = Value
'        End Set
'    End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Modified Person Id")>
'    'Public Property modified_person_id() As Integer?
'    '    Get
'    '        Return Me._modified_person_id
'    '    End Get
'    '    Set
'    '        Me._modified_person_id = Value
'    '    End Set
'    'End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Process Stage")>
'    'Public Property process_stage() As String
'    '    Get
'    '        Return Me._process_stage
'    '    End Get
'    '    Set
'    '        Me._process_stage = Value
'    '    End Set
'    'End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Modified Date")>
'    'Public Property modified_date() As DateTime?
'    '    Get
'    '        Return Me._modified_date
'    '    End Get
'    '    Set
'    '        Me._modified_date = Value
'    '    End Set
'    'End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal brrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = brrow

'        Me.bolt_id = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
'            Me.local_bolt_group_id = DBtoNullableInt(dr.Item("local_bolt_group_id"))
'        End If
'        'Me.work_order_seq_num = DBtoNullableDbl(dr.Item("work_order_seq_num"))
'        Me.rating = DBtoNullableDbl(dr.Item("rating")) 'same in all 
'        Me.result_lkup = DBtoStr(dr.Item("result_lkup")) 'same in all
'        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
'        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
'        'Me.modified_date = DBtoStr(dr.Item("modified_date"))

'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""

'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)


'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""

'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_id")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")


'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_id = " & Me.bolt_id.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As BoltResults = TryCast(other, BoltResults)
'        If otherToCompare Is Nothing Then Return False


'    End Function
'#End Region

'End Class

'Partial Public Class StiffenerGroup
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Stiffener Groups"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.stiffeners"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Group (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Group_INSERT
'        SQLInsert = SQLInsert.Replace("[STIFFENER GROUP VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[STIFFENER GROUP FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Bolt Detail
'        If Me.StiffenerDetails.Count > 0 Then
'            SQLInsert = SQLInsert.Replace("--BEGIN --[STIFFENER DETAIL INSERT BEGIN]", "BEGIN --[STIFFENER DETAIL INSERT BEGIN]")
'            SQLInsert = SQLInsert.Replace("--END --[STIFFENER DETAIL INSERT END]", "END --[STIFFENER DETAIL INSERT END]")
'            For Each row As StiffenerDetail In StiffenerDetails
'                SQLInsert = SQLInsert.Replace("--[STIFFENER DETAIL INSERT]", row.SQLInsert)
'            Next
'        End If

'        ''Results
'        'For Each row As StiffenerResults In StiffenerResults
'        '    SQLInsert = SQLInsert.Replace("--BEGIN --[STIFFENER RESULTS INSERT BEGIN]", "BEGIN --[STIFFENER RESULTS INSERT BEGIN]")
'        '    SQLInsert = SQLInsert.Replace("--END --[STIFFENER RESULTS INSERT END]", "END --[STIFFENER RESULTS INSERT END]")
'        '    SQLInsert = SQLInsert.Replace("--[STIFFENER RESULTS INSERT]", row.SQLInsert)
'        'Next

'        Return SQLInsert

'    End Function

'    Public Overrides Function SQLUpdate() As String

'        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Group (UPDATE).sql")
'        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Group_UPDATE
'        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
'        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Stiffener Detail
'        If Me.StiffenerDetails.Count > 0 Then
'            SQLUpdate = SQLUpdate.Replace("--BEGIN --[STIFFENER DETAIL UPDATE BEGIN]", "BEGIN --[STIFFENER DETAIL UPDATE BEGIN]")
'            SQLUpdate = SQLUpdate.Replace("--END --[STIFFENER DETAIL UPDATE END]", "END --[STIFFENER DETAIL UPDATE END]")
'            For Each row As StiffenerDetail In StiffenerDetails
'                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
'                    If IsSomething(row.stiffener_location) Or IsSomething(row.stiffener_width) Or IsSomething(row.stiffener_height) _
'                        Or IsSomething(row.stiffener_thickness) Or IsSomething(row.stiffener_h_notch) Or IsSomething(row.stiffener_v_notch) _
'                        Or IsSomething(row.stiffener_grade) Or IsSomethingString(row.weld_type) Or IsSomething(row.groove_depth) _
'                        Or IsSomething(row.groove_angle) Or IsSomething(row.h_fillet_weld) Or IsSomething(row.v_fillet_weld) _
'                        Or IsSomething(row.weld_strength) Then
'                        SQLUpdate = SQLUpdate.Replace("--[STIFFENER DETAIL INSERT]", row.SQLUpdate)
'                    Else
'                        SQLUpdate = SQLUpdate.Replace("--[STIFFENER DETAIL INSERT]", row.SQLDelete)
'                    End If
'                Else
'                    SQLUpdate = SQLUpdate.Replace("--[STIFFENER DETAIL INSERT]", row.SQLInsert)
'                End If
'            Next
'        End If

'        'For Each row As StiffenerResults In StiffenerResults
'        '    SQLUpdate = SQLUpdate.Replace("--BEGIN --[STIFFENER RESULTS INSERT BEGIN]", "BEGIN --[STIFFENER RESULTS INSERT BEGIN]")
'        '    SQLUpdate = SQLUpdate.Replace("--END --[STIFFENER RESULTS INSERT END]", "END --[STIFFENER RESULTS INSERT END]")
'        '    SQLUpdate = SQLUpdate.Replace("--[STIFFENER RESULTS INSERT]", row.SQLInsert)
'        'Next

'        Return SQLUpdate

'    End Function

'    Public Overrides Function SQLDelete() As String

'        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Group (DELETE).sql")
'        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Group_DELETE
'        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Stiffener Details
'        If Me.StiffenerDetails.Count > 0 Then
'            SQLDelete = SQLDelete.Replace("--BEGIN --[STIFFENER DETAIL DELETE BEGIN]", "BEGIN --[STIFFENER DETAIL DELETE BEGIN]")
'            SQLDelete = SQLDelete.Replace("--END --[STIFFENER DETAIL DELETE END]", "END --[STIFFENER DETAIL DELETE END]")
'            For Each row As StiffenerDetail In StiffenerDetails
'                SQLDelete = SQLDelete.Replace("--[STIFFENER DETAIL INSERT]", row.SQLDelete)
'            Next
'        End If

'        Return SQLDelete

'    End Function

'#End Region

'#Region "Define"
'    Private _local_id As Integer?
'    Private _local_plate_id As Integer?
'    Private _ID As Integer?
'    Private _plate_details_id As Integer?
'    Private _stiffener_name As String

'    Public Property StiffenerDetails As New List(Of StiffenerDetail)
'    'Public Property StiffenerResults As New List(Of StiffenerResults)

'    <Category("Stiffener Groups"), Description(""), DisplayName("Local Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me._local_id
'        End Get
'        Set
'            Me._local_id = Value
'        End Set
'    End Property
'    <Category("Stiffener Groups"), Description(""), DisplayName("Local Plate Id")>
'    Public Property local_plate_id() As Integer?
'        Get
'            Return Me._local_plate_id
'        End Get
'        Set
'            Me._local_plate_id = Value
'        End Set
'    End Property

'    <Category("Stiffener Groups"), Description(""), DisplayName("Id")>
'    Public Property ID() As Integer?
'        Get
'            Return Me._ID
'        End Get
'        Set
'            Me._ID = Value
'        End Set
'    End Property
'    <Category("Stiffener Groups"), Description(""), DisplayName("Plate Details Id")>
'    Public Property plate_details_id() As Integer?
'        Get
'            Return Me._plate_details_id
'        End Get
'        Set
'            Me._plate_details_id = Value
'        End Set
'    End Property
'    <Category("Stiffener Groups"), Description(""), DisplayName("Stiffener Name")>
'    Public Property stiffener_name() As String
'        Get
'            Return Me._stiffener_name
'        End Get
'        Set
'            Me._stiffener_name = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal sgrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = sgrow
'        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_plate_id")))
'        Me.ID = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_id = DBtoNullableInt(dr.Item("local_stiffener_group_id"))
'            Me.local_plate_id = DBtoNullableInt(dr.Item("local_plate_id"))
'        End If
'        'Me.plate_details_id = If(EDStruefalse, DBtoNullableInt(dr.Item("plate_details_id")), DBtoNullableInt(dr.Item("local_plate_id")))
'        If EDStruefalse = True Then 'Only pull in when referencing EDS
'            Me.plate_details_id = DBtoNullableInt(dr.Item("plate_details_id"))
'        End If
'        Me.stiffener_name = If(EDStruefalse, DBtoStr(dr.Item("stiffener_name")), DBtoStr(dr.Item("group_name")))

'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_name.ToString.FormatDBValue)

'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_details_id")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_name")

'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_details_id = " & "@SubLevel2ID")
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_name = " & Me.stiffener_name.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As StiffenerGroup = TryCast(other, StiffenerGroup)
'        If otherToCompare Is Nothing Then Return False

'        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
'        'Equals = If(Me.plate_details_id.CheckChange(otherToCompare.plate_details_id, changes, categoryName, "Plate Details Id"), Equals, False)
'        Equals = If(Me.stiffener_name.CheckChange(otherToCompare.stiffener_name, changes, categoryName, "Stiffener Name"), Equals, False)

'        'Stiffener Details
'        If Me.StiffenerDetails.Count > 0 Then
'            Equals = If(Me.StiffenerDetails.CheckChange(otherToCompare.StiffenerDetails, changes, categoryName, "Stiffener Details"), Equals, False)
'        End If

'    End Function
'#End Region

'End Class

'Partial Public Class StiffenerDetail
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Stiffener Details"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.stiffener_details"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Detail (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Detail_INSERT
'        SQLInsert = SQLInsert.Replace("[STIFFENER DETAIL VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[STIFFENER DETAIL FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLInsert

'    End Function

'    Public Overrides Function SQLUpdate() As String

'        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Detail (UPDATE).sql")
'        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Detail_UPDATE
'        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
'        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLUpdate

'    End Function

'    Public Overrides Function SQLDelete() As String

'        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Detail (DELETE).sql")
'        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Detail_DELETE
'        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLDelete

'    End Function

'#End Region

'#Region "Define"
'    'Private _local_id As Integer?
'    Private _local_plate_id As Integer?
'    Private _local_group_id As Integer?
'    Private _ID As Integer?
'    Private _stiffener_id As Integer?
'    Private _stiffener_location As Double?
'    Private _stiffener_width As Double?
'    Private _stiffener_height As Double?
'    Private _stiffener_thickness As Double?
'    Private _stiffener_h_notch As Double?
'    Private _stiffener_v_notch As Double?
'    Private _stiffener_grade As Double?
'    Private _weld_type As String
'    Private _groove_depth As Double?
'    Private _groove_angle As Double?
'    Private _h_fillet_weld As Double?
'    Private _v_fillet_weld As Double?
'    Private _weld_strength As Double?


'    Public Property CCIplateMaterials As New List(Of CCIplateMaterial)

'    '<Category("Bolt Details"), Description(""), DisplayName("Local Id")>
'    'Public Property local_id() As Integer?
'    '    Get
'    '        Return Me._local_id
'    '    End Get
'    '    Set
'    '        Me._local_id = Value
'    '    End Set
'    'End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Local Plate Id")>
'    Public Property local_plate_id() As Integer?
'        Get
'            Return Me._local_plate_id
'        End Get
'        Set
'            Me._local_plate_id = Value
'        End Set
'    End Property
'    <Category("Bolt Details"), Description(""), DisplayName("Local Group Id")>
'    Public Property local_group_id() As Integer?
'        Get
'            Return Me._local_group_id
'        End Get
'        Set
'            Me._local_group_id = Value
'        End Set
'    End Property

'    <Category("Stiffener Details"), Description(""), DisplayName("Id")>
'    Public Property ID() As Integer?
'        Get
'            Return Me._ID
'        End Get
'        Set
'            Me._ID = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Id")>
'    Public Property stiffener_id() As Integer?
'        Get
'            Return Me._stiffener_id
'        End Get
'        Set
'            Me._stiffener_id = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Location")>
'    Public Property stiffener_location() As Double?
'        Get
'            Return Me._stiffener_location
'        End Get
'        Set
'            Me._stiffener_location = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Width")>
'    Public Property stiffener_width() As Double?
'        Get
'            Return Me._stiffener_width
'        End Get
'        Set
'            Me._stiffener_width = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Height")>
'    Public Property stiffener_height() As Double?
'        Get
'            Return Me._stiffener_height
'        End Get
'        Set
'            Me._stiffener_height = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Thickness")>
'    Public Property stiffener_thickness() As Double?
'        Get
'            Return Me._stiffener_thickness
'        End Get
'        Set
'            Me._stiffener_thickness = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener H Notch")>
'    Public Property stiffener_h_notch() As Double?
'        Get
'            Return Me._stiffener_h_notch
'        End Get
'        Set
'            Me._stiffener_h_notch = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener V Notch")>
'    Public Property stiffener_v_notch() As Double?
'        Get
'            Return Me._stiffener_v_notch
'        End Get
'        Set
'            Me._stiffener_v_notch = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Grade")>
'    Public Property stiffener_grade() As Double?
'        Get
'            Return Me._stiffener_grade
'        End Get
'        Set
'            Me._stiffener_grade = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Weld Type")>
'    Public Property weld_type() As String
'        Get
'            Return Me._weld_type
'        End Get
'        Set
'            Me._weld_type = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Groove Depth")>
'    Public Property groove_depth() As Double?
'        Get
'            Return Me._groove_depth
'        End Get
'        Set
'            Me._groove_depth = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Groove Angle")>
'    Public Property groove_angle() As Double?
'        Get
'            Return Me._groove_angle
'        End Get
'        Set
'            Me._groove_angle = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("H Fillet Weld")>
'    Public Property h_fillet_weld() As Double?
'        Get
'            Return Me._h_fillet_weld
'        End Get
'        Set
'            Me._h_fillet_weld = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("V Fillet Weld")>
'    Public Property v_fillet_weld() As Double?
'        Get
'            Return Me._v_fillet_weld
'        End Get
'        Set
'            Me._v_fillet_weld = Value
'        End Set
'    End Property
'    <Category("Stiffener Details"), Description(""), DisplayName("Weld Strength")>
'    Public Property weld_strength() As Double?
'        Get
'            Return Me._weld_strength
'        End Get
'        Set
'            Me._weld_strength = Value
'        End Set
'    End Property



'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal bdrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = bdrow
'        Me.ID = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            'Me.local_id = DBtoNullableInt(dr.Item("local_id")) 'nothing references stiffener details so deactivating
'            Me.local_plate_id = DBtoNullableInt(dr.Item("local_plate_id"))
'            Me.local_group_id = DBtoNullableInt(dr.Item("local_group_id"))
'        End If
'        Me.stiffener_id = If(EDStruefalse, DBtoNullableInt(dr.Item("stiffener_id")), DBtoNullableInt(dr.Item("group_id")))
'        Me.stiffener_location = DBtoNullableDbl(dr.Item("stiffener_location"))
'        Me.stiffener_width = DBtoNullableDbl(dr.Item("stiffener_width"))
'        Me.stiffener_height = DBtoNullableDbl(dr.Item("stiffener_height"))
'        Me.stiffener_thickness = DBtoNullableDbl(dr.Item("stiffener_thickness"))
'        Me.stiffener_h_notch = DBtoNullableDbl(dr.Item("stiffener_h_notch"))
'        Me.stiffener_v_notch = DBtoNullableDbl(dr.Item("stiffener_v_notch"))
'        Me.stiffener_grade = DBtoNullableDbl(dr.Item("stiffener_grade"))
'        Me.weld_type = DBtoStr(dr.Item("weld_type"))
'        Me.groove_depth = DBtoNullableDbl(dr.Item("groove_depth"))
'        Me.groove_angle = DBtoNullableDbl(dr.Item("groove_angle"))
'        Me.h_fillet_weld = DBtoNullableDbl(dr.Item("h_fillet_weld"))
'        Me.v_fillet_weld = DBtoNullableDbl(dr.Item("v_fillet_weld"))
'        Me.weld_strength = DBtoNullableDbl(dr.Item("weld_strength"))

'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_id.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_location.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_width.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_height.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_thickness.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_h_notch.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_v_notch.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_grade.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_type.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.groove_depth.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.groove_angle.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.h_fillet_weld.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.v_fillet_weld.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_strength.ToString.FormatDBValue)

'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_id")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_location")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_width")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_height")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_thickness")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_h_notch")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_v_notch")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_grade")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_type")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("groove_depth")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("groove_angle")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("h_fillet_weld")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("v_fillet_weld")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_strength")

'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_id = " & "@SubLevel3ID")
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_location = " & Me.stiffener_location.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_width = " & Me.stiffener_width.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_height = " & Me.stiffener_height.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_thickness = " & Me.stiffener_thickness.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_h_notch = " & Me.stiffener_h_notch.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_v_notch = " & Me.stiffener_v_notch.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_grade = " & Me.stiffener_grade.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_type = " & Me.weld_type.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("groove_depth = " & Me.groove_depth.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("groove_angle = " & Me.groove_angle.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("h_fillet_weld = " & Me.h_fillet_weld.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("v_fillet_weld = " & Me.v_fillet_weld.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_strength = " & Me.weld_strength.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As StiffenerDetail = TryCast(other, StiffenerDetail)
'        If otherToCompare Is Nothing Then Return False

'        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
'        Equals = If(Me.stiffener_id.CheckChange(otherToCompare.stiffener_id, changes, categoryName, "Stiffener Id"), Equals, False)
'        Equals = If(Me.stiffener_location.CheckChange(otherToCompare.stiffener_location, changes, categoryName, "Stiffener Location"), Equals, False)
'        Equals = If(Me.stiffener_width.CheckChange(otherToCompare.stiffener_width, changes, categoryName, "Stiffener Width"), Equals, False)
'        Equals = If(Me.stiffener_height.CheckChange(otherToCompare.stiffener_height, changes, categoryName, "Stiffener Height"), Equals, False)
'        Equals = If(Me.stiffener_thickness.CheckChange(otherToCompare.stiffener_thickness, changes, categoryName, "Stiffener Thickness"), Equals, False)
'        Equals = If(Me.stiffener_h_notch.CheckChange(otherToCompare.stiffener_h_notch, changes, categoryName, "Stiffener H Notch"), Equals, False)
'        Equals = If(Me.stiffener_v_notch.CheckChange(otherToCompare.stiffener_v_notch, changes, categoryName, "Stiffener V Notch"), Equals, False)
'        Equals = If(Me.stiffener_grade.CheckChange(otherToCompare.stiffener_grade, changes, categoryName, "Stiffener Grade"), Equals, False)
'        Equals = If(Me.weld_type.CheckChange(otherToCompare.weld_type, changes, categoryName, "Weld Type"), Equals, False)
'        Equals = If(Me.groove_depth.CheckChange(otherToCompare.groove_depth, changes, categoryName, "Groove Depth"), Equals, False)
'        Equals = If(Me.groove_angle.CheckChange(otherToCompare.groove_angle, changes, categoryName, "Groove Angle"), Equals, False)
'        Equals = If(Me.h_fillet_weld.CheckChange(otherToCompare.h_fillet_weld, changes, categoryName, "H Fillet Weld"), Equals, False)
'        Equals = If(Me.v_fillet_weld.CheckChange(otherToCompare.v_fillet_weld, changes, categoryName, "V Fillet Weld"), Equals, False)
'        Equals = If(Me.weld_strength.CheckChange(otherToCompare.weld_strength, changes, categoryName, "Weld Strength"), Equals, False)

'    End Function
'#End Region

'End Class

'Partial Public Class StiffenerResults
'    Inherits EDSObjectWithQueries
'    'StiffenerResults are currently not being referenced. Stiffeners reported with plate details. 
'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Stiffener Results"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.stiffener_results"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Result (INSERT).sql")
'        'SQLInsert = CCI_Engineering_Templates.My.Resources.Stiffener_Result_INSERT
'        SQLInsert = SQLInsert.Replace("[STIFFENER RESULT VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[STIFFENER RESULT FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLInsert

'    End Function

'    'Public Overrides Function SQLUpdate() As String

'    '    SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Connection Material (UPDATE).sql")
'    '    SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'    '    SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
'    '    SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

'    '    Return SQLUpdate

'    'End Function

'    'Public Overrides Function SQLDelete() As String

'    '    SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Connection Material (DELETE).sql")
'    '    SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'    '    SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

'    '    Return SQLDelete

'    'End Function

'#End Region

'#Region "Define"
'    Private _stiffener_id As Integer?
'    Private _local_id As Integer?
'    Private _local_stiffener_group_id As Integer?
'    'Private _work_order_seq_num As Double? 'not provided in Excel
'    Private _rating As Double?
'    Private _result_lkup As String
'    'Private _modified_person_id As Integer? 'not provided in Excel
'    'Private _process_stage As String 'not provided in Excel
'    'Private _modified_date As DateTime? 'not provided in Excel

'    <Category("Stiffener Results"), Description(""), DisplayName("Stiffener Id")>
'    Public Property stiffener_id() As Integer?
'        Get
'            Return Me._stiffener_id
'        End Get
'        Set
'            Me._stiffener_id = Value
'        End Set
'    End Property
'    <Category("Stiffener Results"), Description(""), DisplayName("Local Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me._local_id
'        End Get
'        Set
'            Me._local_id = Value
'        End Set
'    End Property
'    <Category("Stiffener Results"), Description(""), DisplayName("Local Bolt Group Id")>
'    Public Property local_stiffener_group_id() As Integer?
'        Get
'            Return Me._local_stiffener_group_id
'        End Get
'        Set
'            Me._local_stiffener_group_id = Value
'        End Set
'    End Property
'    '<Category("Stiffener Results"), Description(""), DisplayName("Work Order Seq Num")>
'    'Public Property work_order_seq_num() As Double?
'    '    Get
'    '        Return Me._work_order_seq_num
'    '    End Get
'    '    Set
'    '        Me._work_order_seq_num = Value
'    '    End Set
'    'End Property
'    <Category("Stiffener Results"), Description(""), DisplayName("Rating")>
'    Public Property rating() As Double?
'        Get
'            Return Me._rating
'        End Get
'        Set
'            Me._rating = Value
'        End Set
'    End Property
'    <Category("Stiffener Results"), Description(""), DisplayName("Result Lkup")>
'    Public Property result_lkup() As String
'        Get
'            Return Me._result_lkup
'        End Get
'        Set
'            Me._result_lkup = Value
'        End Set
'    End Property
'    '<Category("Stiffener Results"), Description(""), DisplayName("Modified Person Id")>
'    'Public Property modified_person_id() As Integer?
'    '    Get
'    '        Return Me._modified_person_id
'    '    End Get
'    '    Set
'    '        Me._modified_person_id = Value
'    '    End Set
'    'End Property
'    '<Category("Stiffener Results"), Description(""), DisplayName("Process Stage")>
'    'Public Property process_stage() As String
'    '    Get
'    '        Return Me._process_stage
'    '    End Get
'    '    Set
'    '        Me._process_stage = Value
'    '    End Set
'    'End Property
'    '<Category("Stiffener Results"), Description(""), DisplayName("Modified Date")>
'    'Public Property modified_date() As DateTime?
'    '    Get
'    '        Return Me._modified_date
'    '    End Get
'    '    Set
'    '        Me._modified_date = Value
'    '    End Set
'    'End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal brrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = brrow

'        Me.stiffener_id = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_id = DBtoNullableInt(dr.Item("local_id"))
'            Me.local_stiffener_group_id = DBtoNullableInt(dr.Item("local_bolt_group_id"))
'        End If
'        'Me.work_order_seq_num = DBtoNullableDbl(dr.Item("work_order_seq_num"))
'        Me.rating = DBtoNullableDbl(dr.Item("rating")) 'same in all 
'        Me.result_lkup = DBtoStr(dr.Item("result_lkup")) 'same in all
'        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
'        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
'        'Me.modified_date = DBtoStr(dr.Item("modified_date"))

'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""

'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)


'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""

'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_id")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")


'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_id = " & Me.stiffener_id.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As StiffenerResults = TryCast(other, StiffenerResults)
'        If otherToCompare Is Nothing Then Return False


'    End Function
'#End Region

'End Class

'Partial Public Class BridgeStiffenerDetail
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Bridge Stiffener Details"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.bridge_stiffeners"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Bridge Stiffener Detail (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Bridge_Stiffener_Detail_INSERT
'        SQLInsert = SQLInsert.Replace("[BRIDGE STIFFENER DETAIL VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[BRIDGE STIFFENER DETAIL FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Plate Material
'        For Each row As CCIplateMaterial In CCIplateMaterials
'            SQLInsert = SQLInsert.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert)
'        Next

'        ''Results - Probably placing under connections
'        'For Each row As BoltResults In BoltResults
'        '    SQLInsert = SQLInsert.Replace("--BEGIN --[BOLT RESULTS INSERT BEGIN]", "BEGIN --[BOLT RESULTS INSERT BEGIN]")
'        '    SQLInsert = SQLInsert.Replace("--END --[BOLT RESULTS INSERT END]", "END --[BOLT RESULTS INSERT END]")
'        '    SQLInsert = SQLInsert.Replace("--[BOLT RESULTS INSERT]", row.SQLInsert)
'        'Next

'        Return SQLInsert

'    End Function

'    Public Overrides Function SQLUpdate() As String

'        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Bridge Stiffener Detail (UPDATE).sql")
'        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Bridge_Stiffener_Detail_UPDATE
'        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
'        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

'        'Plate Material
'        For Each row As CCIplateMaterial In CCIplateMaterials
'            SQLUpdate = SQLUpdate.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert) 'Can only insert materials, no deleting or updating since database is referenced by all BUs. 
'        Next

'        Return SQLUpdate

'    End Function

'    Public Overrides Function SQLDelete() As String

'        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Bridge Stiffener Detail (DELETE).sql")
'        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Bridge_Stiffener_Detail_DELETE
'        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
'        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLDelete

'    End Function

'#End Region

'#Region "Define"
'    Private _local_id As Integer?
'    Private _local_connection_id As Integer?
'    Private _ID As Integer?
'    Private _connection_id As Integer?
'    Private _stiffener_type As String
'    Private _analysis_type As String
'    Private _quantity As Double?
'    Private _bridge_stiffener_width As Double?
'    Private _bridge_stiffener_thickness As Double?
'    Private _bridge_stiffener_material As Integer?
'    Private _unbraced_length As Double?
'    Private _total_length As Double?
'    Private _weld_size As Double?
'    Private _exx As Double?
'    Private _upper_weld_length As Double?
'    Private _lower_weld_length As Double?
'    Private _upper_plate_width As Double?
'    Private _lower_plate_width As Double?
'    Private _neglect_flange_connection As Boolean?
'    Private _bolt_hole_diameter As Double?
'    Private _bolt_qty_eccentric As Double?
'    Private _bolt_qty_shear As Double?
'    Private _intermediate_bolt_spacing As Double?
'    Private _bolt_diameter As Double?
'    Private _bolt_sleeve_diameter As Double?
'    Private _washer_diameter As Double?
'    Private _bolt_tensile_strength As Double?
'    Private _bolt_allowable_shear As Double?
'    Private _exx_shim_plate As Double?
'    Private _filler_shim_thickness As Double?

'    Public Property CCIplateMaterials As New List(Of CCIplateMaterial)

'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Local Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me._local_id
'        End Get
'        Set
'            Me._local_id = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Local Connection Id")>
'    Public Property local_connection_id() As Integer?
'        Get
'            Return Me._local_connection_id
'        End Get
'        Set
'            Me._local_connection_id = Value
'        End Set
'    End Property

'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Id")>
'    Public Property ID() As Integer?
'        Get
'            Return Me._ID
'        End Get
'        Set
'            Me._ID = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Connection Id")>
'    Public Property connection_id() As Integer?
'        Get
'            Return Me._connection_id
'        End Get
'        Set
'            Me._connection_id = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Stiffener Type")>
'    Public Property stiffener_type() As String
'        Get
'            Return Me._stiffener_type
'        End Get
'        Set
'            Me._stiffener_type = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Analysis Type")>
'    Public Property analysis_type() As String
'        Get
'            Return Me._analysis_type
'        End Get
'        Set
'            Me._analysis_type = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Quantity")>
'    Public Property quantity() As Double?
'        Get
'            Return Me._quantity
'        End Get
'        Set
'            Me._quantity = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bridge Stiffener Width")>
'    Public Property bridge_stiffener_width() As Double?
'        Get
'            Return Me._bridge_stiffener_width
'        End Get
'        Set
'            Me._bridge_stiffener_width = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bridge Stiffener Thickness")>
'    Public Property bridge_stiffener_thickness() As Double?
'        Get
'            Return Me._bridge_stiffener_thickness
'        End Get
'        Set
'            Me._bridge_stiffener_thickness = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bridge Stiffener Material")>
'    Public Property bridge_stiffener_material() As Integer?
'        Get
'            Return Me._bridge_stiffener_material
'        End Get
'        Set
'            Me._bridge_stiffener_material = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Unbraced Length")>
'    Public Property unbraced_length() As Double?
'        Get
'            Return Me._unbraced_length
'        End Get
'        Set
'            Me._unbraced_length = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Total Length")>
'    Public Property total_length() As Double?
'        Get
'            Return Me._total_length
'        End Get
'        Set
'            Me._total_length = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Weld Size")>
'    Public Property weld_size() As Double?
'        Get
'            Return Me._weld_size
'        End Get
'        Set
'            Me._weld_size = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Exx")>
'    Public Property exx() As Double?
'        Get
'            Return Me._exx
'        End Get
'        Set
'            Me._exx = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Upper Weld Length")>
'    Public Property upper_weld_length() As Double?
'        Get
'            Return Me._upper_weld_length
'        End Get
'        Set
'            Me._upper_weld_length = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Lower Weld Length")>
'    Public Property lower_weld_length() As Double?
'        Get
'            Return Me._lower_weld_length
'        End Get
'        Set
'            Me._lower_weld_length = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Upper Plate Width")>
'    Public Property upper_plate_width() As Double?
'        Get
'            Return Me._upper_plate_width
'        End Get
'        Set
'            Me._upper_plate_width = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Lower Plate Width")>
'    Public Property lower_plate_width() As Double?
'        Get
'            Return Me._lower_plate_width
'        End Get
'        Set
'            Me._lower_plate_width = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Neglect Flange Connection")>
'    Public Property neglect_flange_connection() As Boolean?
'        Get
'            Return Me._neglect_flange_connection
'        End Get
'        Set
'            Me._neglect_flange_connection = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Hole Diameter")>
'    Public Property bolt_hole_diameter() As Double?
'        Get
'            Return Me._bolt_hole_diameter
'        End Get
'        Set
'            Me._bolt_hole_diameter = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Qty Eccentric")>
'    Public Property bolt_qty_eccentric() As Double?
'        Get
'            Return Me._bolt_qty_eccentric
'        End Get
'        Set
'            Me._bolt_qty_eccentric = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Qty Shear")>
'    Public Property bolt_qty_shear() As Double?
'        Get
'            Return Me._bolt_qty_shear
'        End Get
'        Set
'            Me._bolt_qty_shear = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Intermediate Bolt Spacing")>
'    Public Property intermediate_bolt_spacing() As Double?
'        Get
'            Return Me._intermediate_bolt_spacing
'        End Get
'        Set
'            Me._intermediate_bolt_spacing = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Diameter")>
'    Public Property bolt_diameter() As Double?
'        Get
'            Return Me._bolt_diameter
'        End Get
'        Set
'            Me._bolt_diameter = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Sleeve Diameter")>
'    Public Property bolt_sleeve_diameter() As Double?
'        Get
'            Return Me._bolt_sleeve_diameter
'        End Get
'        Set
'            Me._bolt_sleeve_diameter = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Washer Diameter")>
'    Public Property washer_diameter() As Double?
'        Get
'            Return Me._washer_diameter
'        End Get
'        Set
'            Me._washer_diameter = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Tensile Strength")>
'    Public Property bolt_tensile_strength() As Double?
'        Get
'            Return Me._bolt_tensile_strength
'        End Get
'        Set
'            Me._bolt_tensile_strength = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Allowable Shear")>
'    Public Property bolt_allowable_shear() As Double?
'        Get
'            Return Me._bolt_allowable_shear
'        End Get
'        Set
'            Me._bolt_allowable_shear = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Exx Shim Plate")>
'    Public Property exx_shim_plate() As Double?
'        Get
'            Return Me._exx_shim_plate
'        End Get
'        Set
'            Me._exx_shim_plate = Value
'        End Set
'    End Property
'    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Filler Shim Thickness")>
'    Public Property filler_shim_thickness() As Double?
'        Get
'            Return Me._filler_shim_thickness
'        End Get
'        Set
'            Me._filler_shim_thickness = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal bsdrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = bsdrow
'        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_plate_id")))
'        Me.ID = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_id = DBtoNullableInt(dr.Item("local_bridge_stiffener_id")) 'currently not being referenced for anything
'        End If
'        Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id")) 'need to store local_connection_id within EDS since user may adjust relationship to any elevation and need to identify this as a change to perform update function. 
'        Me.connection_id = If(EDStruefalse, DBtoNullableInt(dr.Item("plate_id")), DBtoNullableInt(dr.Item("connection_id")))
'        'If EDStruefalse = True Then 'Only pull in when referencing EDS
'        '    Me.plate_id = DBtoNullableInt(dr.Item("plate_id"))
'        'End If
'        Me.stiffener_type = DBtoStr(dr.Item("stiffener_type"))
'        Me.analysis_type = DBtoStr(dr.Item("analysis_type"))
'        Me.quantity = DBtoNullableDbl(dr.Item("quantity"))
'        Me.bridge_stiffener_width = DBtoNullableDbl(dr.Item("bridge_stiffener_width"))
'        Me.bridge_stiffener_thickness = DBtoNullableDbl(dr.Item("bridge_stiffener_thickness"))
'        Me.bridge_stiffener_material = DBtoNullableInt(dr.Item("bridge_stiffener_material"))
'        Me.unbraced_length = DBtoNullableDbl(dr.Item("unbraced_length"))
'        Me.total_length = DBtoNullableDbl(dr.Item("total_length"))
'        Me.weld_size = DBtoNullableDbl(dr.Item("weld_size"))
'        Me.exx = DBtoNullableDbl(dr.Item("exx"))
'        Me.upper_weld_length = DBtoNullableDbl(dr.Item("upper_weld_length"))
'        Me.lower_weld_length = DBtoNullableDbl(dr.Item("lower_weld_length"))
'        Me.upper_plate_width = DBtoNullableDbl(dr.Item("upper_plate_width"))
'        Me.lower_plate_width = DBtoNullableDbl(dr.Item("lower_plate_width"))
'        Me.neglect_flange_connection = If(EDStruefalse, DBtoNullableBool(dr.Item("neglect_flange_connection")), If(DBtoStr(dr.Item("neglect_flange_connection")) = "Yes", True, If(DBtoStr(dr.Item("neglect_flange_connection")) = "No", False, DBtoNullableBool(dr.Item("neglect_flange_connection")))))
'        Me.bolt_hole_diameter = DBtoNullableDbl(dr.Item("bolt_hole_diameter"))
'        Me.bolt_qty_eccentric = DBtoNullableDbl(dr.Item("bolt_qty_eccentric"))
'        Me.bolt_qty_shear = DBtoNullableDbl(dr.Item("bolt_qty_shear"))
'        Me.intermediate_bolt_spacing = DBtoNullableDbl(dr.Item("intermediate_bolt_spacing"))
'        Me.bolt_diameter = DBtoNullableDbl(dr.Item("bolt_diameter"))
'        Me.bolt_sleeve_diameter = DBtoNullableDbl(dr.Item("bolt_sleeve_diameter"))
'        Me.washer_diameter = DBtoNullableDbl(dr.Item("washer_diameter"))
'        Me.bolt_tensile_strength = DBtoNullableDbl(dr.Item("bolt_tensile_strength"))
'        Me.bolt_allowable_shear = DBtoNullableDbl(dr.Item("bolt_allowable_shear"))
'        Me.exx_shim_plate = DBtoNullableDbl(dr.Item("exx_shim_plate"))
'        Me.filler_shim_thickness = DBtoNullableDbl(dr.Item("filler_shim_thickness"))

'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_connection_id.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_id.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_type.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.analysis_type.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.quantity.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bridge_stiffener_width.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bridge_stiffener_thickness.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bridge_stiffener_material.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.unbraced_length.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.total_length.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_size.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.exx.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.upper_weld_length.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lower_weld_length.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.upper_plate_width.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lower_plate_width.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.neglect_flange_connection.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_hole_diameter.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_qty_eccentric.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_qty_shear.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.intermediate_bolt_spacing.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_diameter.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_sleeve_diameter.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.washer_diameter.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_tensile_strength.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_allowable_shear.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.exx_shim_plate.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.filler_shim_thickness.ToString.FormatDBValue)

'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""
'        SQLInsertFields = SQLInsertFields.AddtoDBString("local_connection_id")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_id")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_type")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("analysis_type")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("quantity")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bridge_stiffener_width")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bridge_stiffener_thickness")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bridge_stiffener_material")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("unbraced_length")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("total_length")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_size")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("exx")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("upper_weld_length")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("lower_weld_length")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("upper_plate_width")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("lower_plate_width")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("neglect_flange_connection")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_hole_diameter")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_qty_eccentric")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_qty_shear")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("intermediate_bolt_spacing")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_diameter")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_sleeve_diameter")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("washer_diameter")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_tensile_strength")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_allowable_shear")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("exx_shim_plate")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("filler_shim_thickness")

'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_connection_id = " & Me.local_connection_id.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_id = " & "@SubLevel1ID")
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_type = " & Me.stiffener_type.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("analysis_type = " & Me.analysis_type.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("quantity = " & Me.quantity.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bridge_stiffener_width = " & Me.bridge_stiffener_width.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bridge_stiffener_thickness = " & Me.bridge_stiffener_thickness.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bridge_stiffener_material = " & Me.bridge_stiffener_material.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bridge_stiffener_material = " & "@SubLevel3ID")
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("unbraced_length = " & Me.unbraced_length.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("total_length = " & Me.total_length.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_size = " & Me.weld_size.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("exx = " & Me.exx.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("upper_weld_length = " & Me.upper_weld_length.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lower_weld_length = " & Me.lower_weld_length.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("upper_plate_width = " & Me.upper_plate_width.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lower_plate_width = " & Me.lower_plate_width.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("neglect_flange_connection = " & Me.neglect_flange_connection.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_hole_diameter = " & Me.bolt_hole_diameter.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_qty_eccentric = " & Me.bolt_qty_eccentric.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_qty_shear = " & Me.bolt_qty_shear.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("intermediate_bolt_spacing = " & Me.intermediate_bolt_spacing.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_diameter = " & Me.bolt_diameter.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_sleeve_diameter = " & Me.bolt_sleeve_diameter.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("washer_diameter = " & Me.washer_diameter.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_tensile_strength = " & Me.bolt_tensile_strength.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_allowable_shear = " & Me.bolt_allowable_shear.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("exx_shim_plate = " & Me.exx_shim_plate.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("filler_shim_thickness = " & Me.filler_shim_thickness.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'plate_material references the local id when coming from Excel. Need to convert to EDS ID when performing Equals function
'        Dim material As Integer?
'        For Each row As CCIplateMaterial In CCIplateMaterials
'            If Me.bridge_stiffener_material = row.local_id And row.ID > 0 Then
'                material = row.ID
'                Exit For
'            End If
'        Next

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As BridgeStiffenerDetail = TryCast(other, BridgeStiffenerDetail)
'        If otherToCompare Is Nothing Then Return False

'        Equals = If(Me.local_connection_id.CheckChange(otherToCompare.local_connection_id, changes, categoryName, "Plate Id"), Equals, False)
'        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
'        'Equals = If(Me.plate_id.CheckChange(otherToCompare.plate_id, changes, categoryName, "Plate Id"), Equals, False)
'        Equals = If(Me.stiffener_type.CheckChange(otherToCompare.stiffener_type, changes, categoryName, "Stiffener Type"), Equals, False)
'        Equals = If(Me.analysis_type.CheckChange(otherToCompare.analysis_type, changes, categoryName, "Analysis Type"), Equals, False)
'        Equals = If(Me.quantity.CheckChange(otherToCompare.quantity, changes, categoryName, "Quantity"), Equals, False)
'        Equals = If(Me.bridge_stiffener_width.CheckChange(otherToCompare.bridge_stiffener_width, changes, categoryName, "Bridge Stiffener Width"), Equals, False)
'        Equals = If(Me.bridge_stiffener_thickness.CheckChange(otherToCompare.bridge_stiffener_thickness, changes, categoryName, "Bridge Stiffener Thickness"), Equals, False)
'        'Equals = If(Me.bridge_stiffener_material.CheckChange(otherToCompare.bridge_stiffener_material, changes, categoryName, "Bridge Stiffener Material"), Equals, False)
'        Equals = If(material.CheckChange(otherToCompare.bridge_stiffener_material, changes, categoryName, "Bridge Stiffener Material"), Equals, False)
'        Equals = If(Me.unbraced_length.CheckChange(otherToCompare.unbraced_length, changes, categoryName, "Unbraced Length"), Equals, False)
'        Equals = If(Me.total_length.CheckChange(otherToCompare.total_length, changes, categoryName, "Total Length"), Equals, False)
'        Equals = If(Me.weld_size.CheckChange(otherToCompare.weld_size, changes, categoryName, "Weld Size"), Equals, False)
'        Equals = If(Me.exx.CheckChange(otherToCompare.exx, changes, categoryName, "Exx"), Equals, False)
'        Equals = If(Me.upper_weld_length.CheckChange(otherToCompare.upper_weld_length, changes, categoryName, "Upper Weld Length"), Equals, False)
'        Equals = If(Me.lower_weld_length.CheckChange(otherToCompare.lower_weld_length, changes, categoryName, "Lower Weld Length"), Equals, False)
'        Equals = If(Me.upper_plate_width.CheckChange(otherToCompare.upper_plate_width, changes, categoryName, "Upper Plate Width"), Equals, False)
'        Equals = If(Me.lower_plate_width.CheckChange(otherToCompare.lower_plate_width, changes, categoryName, "Lower Plate Width"), Equals, False)
'        Equals = If(Me.neglect_flange_connection.CheckChange(otherToCompare.neglect_flange_connection, changes, categoryName, "Neglect Flange Connection"), Equals, False)
'        Equals = If(Me.bolt_hole_diameter.CheckChange(otherToCompare.bolt_hole_diameter, changes, categoryName, "Bolt Hole Diameter"), Equals, False)
'        Equals = If(Me.bolt_qty_eccentric.CheckChange(otherToCompare.bolt_qty_eccentric, changes, categoryName, "Bolt Qty Eccentric"), Equals, False)
'        Equals = If(Me.bolt_qty_shear.CheckChange(otherToCompare.bolt_qty_shear, changes, categoryName, "Bolt Qty Shear"), Equals, False)
'        Equals = If(Me.intermediate_bolt_spacing.CheckChange(otherToCompare.intermediate_bolt_spacing, changes, categoryName, "Intermediate Bolt Spacing"), Equals, False)
'        Equals = If(Me.bolt_diameter.CheckChange(otherToCompare.bolt_diameter, changes, categoryName, "Bolt Diameter"), Equals, False)
'        Equals = If(Me.bolt_sleeve_diameter.CheckChange(otherToCompare.bolt_sleeve_diameter, changes, categoryName, "Bolt Sleeve Diameter"), Equals, False)
'        Equals = If(Me.washer_diameter.CheckChange(otherToCompare.washer_diameter, changes, categoryName, "Washer Diameter"), Equals, False)
'        Equals = If(Me.bolt_tensile_strength.CheckChange(otherToCompare.bolt_tensile_strength, changes, categoryName, "Bolt Tensile Strength"), Equals, False)
'        Equals = If(Me.bolt_allowable_shear.CheckChange(otherToCompare.bolt_allowable_shear, changes, categoryName, "Bolt Allowable Shear"), Equals, False)
'        Equals = If(Me.exx_shim_plate.CheckChange(otherToCompare.exx_shim_plate, changes, categoryName, "Exx Shim Plate"), Equals, False)
'        Equals = If(Me.filler_shim_thickness.CheckChange(otherToCompare.filler_shim_thickness, changes, categoryName, "Filler Shim Thickness"), Equals, False)

'        'Materials
'        If Me.CCIplateMaterials.Count > 0 Then
'            Equals = If(Me.CCIplateMaterials.CheckChange(otherToCompare.CCIplateMaterials, changes, categoryName, "CCIplate Materials"), Equals, False)
'        End If

'    End Function
'#End Region

'End Class

'Partial Public Class ConnectionResults
'    Inherits EDSObjectWithQueries

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Connection Results"
'    Public Overrides ReadOnly Property EDSTableName As String = "conn.connection_results"

'    Public Overrides Function SQLInsert() As String

'        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Connection Result (INSERT).sql")
'        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Connection_Result_INSERT
'        SQLInsert = SQLInsert.Replace("[CONNECTION RESULT VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[CONNECTION RESULT FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLInsert

'    End Function


'#End Region

'#Region "Define"
'    Private _plate_id As Integer?
'    Private _local_connection_id As Integer?
'    'Private _local_bolt_group_id As Integer?
'    'Private _work_order_seq_num As Double? 'not provided in Excel
'    Private _rating As Double?
'    Private _result_lkup As String
'    'Private _modified_person_id As Integer? 'not provided in Excel
'    'Private _process_stage As String 'not provided in Excel
'    'Private _modified_date As DateTime? 'not provided in Excel

'    <Category("Connection Results"), Description(""), DisplayName("Plate Id")>
'    Public Property plate_id() As Integer?
'        Get
'            Return Me._plate_id
'        End Get
'        Set
'            Me._plate_id = Value
'        End Set
'    End Property
'    <Category("Connection Results"), Description(""), DisplayName("Local Connection Id")>
'    Public Property local_connection_id() As Integer?
'        Get
'            Return Me._local_connection_id
'        End Get
'        Set
'            Me._local_connection_id = Value
'        End Set
'    End Property
'    '<Category("Connection Results"), Description(""), DisplayName("Local Bolt Group Id")>
'    'Public Property local_bolt_group_id() As Integer?
'    '    Get
'    '        Return Me._local_bolt_group_id
'    '    End Get
'    '    Set
'    '        Me._local_bolt_group_id = Value
'    '    End Set
'    'End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Work Order Seq Num")>
'    'Public Property work_order_seq_num() As Double?
'    '    Get
'    '        Return Me._work_order_seq_num
'    '    End Get
'    '    Set
'    '        Me._work_order_seq_num = Value
'    '    End Set
'    'End Property
'    <Category("Connection Results"), Description(""), DisplayName("Rating")>
'    Public Property rating() As Double?
'        Get
'            Return Me._rating
'        End Get
'        Set
'            Me._rating = Value
'        End Set
'    End Property
'    <Category("Connection Results"), Description(""), DisplayName("Result Lkup")>
'    Public Property result_lkup() As String
'        Get
'            Return Me._result_lkup
'        End Get
'        Set
'            Me._result_lkup = Value
'        End Set
'    End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Modified Person Id")>
'    'Public Property modified_person_id() As Integer?
'    '    Get
'    '        Return Me._modified_person_id
'    '    End Get
'    '    Set
'    '        Me._modified_person_id = Value
'    '    End Set
'    'End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Process Stage")>
'    'Public Property process_stage() As String
'    '    Get
'    '        Return Me._process_stage
'    '    End Get
'    '    Set
'    '        Me._process_stage = Value
'    '    End Set
'    'End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Modified Date")>
'    'Public Property modified_date() As DateTime?
'    '    Get
'    '        Return Me._modified_date
'    '    End Get
'    '    Set
'    '        Me._modified_date = Value
'    '    End Set
'    'End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal crrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = crrow

'        Me.plate_id = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
'        End If
'        'Me.work_order_seq_num = DBtoNullableDbl(dr.Item("work_order_seq_num"))
'        Me.rating = DBtoNullableDbl(dr.Item("rating")) 'same in all 
'        Me.result_lkup = DBtoStr(dr.Item("result_lkup")) 'same in all
'        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
'        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
'        'Me.modified_date = DBtoStr(dr.Item("modified_date"))

'    End Sub

'#End Region

'#Region "Save to EDS"
'    Public Overrides Function SQLInsertValues() As String
'        SQLInsertValues = ""

'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
'        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)


'        Return SQLInsertValues
'    End Function

'    Public Overrides Function SQLInsertFields() As String
'        SQLInsertFields = ""

'        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_id")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
'        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
'        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")


'        Return SQLInsertFields
'    End Function

'    Public Overrides Function SQLUpdateFieldsandValues() As String
'        SQLUpdateFieldsandValues = ""
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_id = " & Me.plate_id.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
'        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

'        Return SQLUpdateFieldsandValues
'    End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As ConnectionResults = TryCast(other, ConnectionResults)
'        If otherToCompare Is Nothing Then Return False


'    End Function
'#End Region

'End Class