Option Strict On

Imports System.ComponentModel
Imports System.Runtime.Serialization
Imports DevExpress.Spreadsheet

''Adding fields to Leg Reinforcement
''1. Excel
''  -Add field to Sub Tables (Sapi) tab
''      -take note of new table range since this needs updated in datatransferer
''  -make sure to update:
''      -Revision History: notes refer to 'internal database'
''      -Import Ranges: add additional column associated to version number and then click 'set document properties'
''2. Datatransferer
''  -Inheritted: increase table range for ExcelDTParams per Details (Sapi) tab
''  -Add associated fields to
''      -Define
''      -Constructor: make sure to add as a try/catch so older sapi version remain compatible
''      -Save to Excel: make sure to handle null values since new field won't include data for anything existing in database
''      -Save to EDS 
''      -Equals
''3.Add SQL Column
''  -Add only to EDS Dev
''  -Save query in the corresponding folder for the current sprint
''  - C:\Users\%username%\Crown Castle USA Inc\ECS - Tools\Database Changes
''      -this will be referenced for updating EDS UAT and EDS PROD

<DataContractAttribute()>
Partial Public Class LegReinforcement
    Inherits EDSExcelObject

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Leg Reinforcement Tool"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "tnx.memb_leg_reinforcement"
        End Get
    End Property
    Public Overrides ReadOnly Property TemplatePath As String
        Get
            Return IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Leg Reinforcement Tool.xlsm")
        End Get
    End Property
    Public Overrides ReadOnly Property Template As Byte()
        Get
            Return CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Tool
        End Get
    End Property
    Public Overrides ReadOnly Property ExcelDTParams As List(Of EXCELDTParameter)
        'Add additional sub table references here. Table names should be consistent with EDS table names. 
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("Leg Reinforcements", "A1:C2", "Details (SAPI)"),
                                            New EXCELDTParameter("Leg Reinforcement Details", "A1:AY201", "Sub Tables (SAPI)"),
                                            New EXCELDTParameter("Leg Reinforcement Results", "A2:D202", "Results (SAPI)")}

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
            'SQLInsert = SQLInsert.Replace("--BEGIN --[LEG REINFORCEMENT DETAIL INSERT BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAIL INSERT BEGIN]")
            'SQLInsert = SQLInsert.Replace("--END --[LEG REINFORCEMENT DETAIL INSERT END]", "END --[LEG REINFORCEMENT DETAIL INSERT END]")
            For Each row As LegReinforcementDetail In LegReinforcementDetails
                If IsSomethingString(row.create_leg) Then
                    SQLInsert = SQLInsert.Replace("--[LEG REINFORCEMENT DETAIL INSERT]", row.SQLInsert)
                    SQLInsert = SQLInsert.Replace("--BEGIN --[LEG REINFORCEMENT DETAIL INSERT BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAIL INSERT BEGIN]")
                    SQLInsert = SQLInsert.Replace("--END --[LEG REINFORCEMENT DETAIL INSERT END]", "END --[LEG REINFORCEMENT DETAIL INSERT END]")
                End If
            Next
        End If

        'note: additional insert commands are imbedded within objects sharing similar relationships

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
            'SQLUpdate = SQLUpdate.Replace("--BEGIN --[LEG REINFORCEMENT DETAIL UPDATE BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAIL UPDATE BEGIN]")
            'SQLUpdate = SQLUpdate.Replace("--END --[LEG REINFORCEMENT DETAIL UPDATE END]", "END --[LEG REINFORCEMENT DETAIL UPDATE END]")
            For Each row As LegReinforcementDetail In LegReinforcementDetails
                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                    'following fields include default values and therefore are removed from check below: end_connection_type, applied_load_type, slenderness_ratio_type, print_bolt_on_connections, reinforcement_type
                    If IsSomethingString(row.create_leg) And (IsSomething(row.leg_load_time_mod_option) Or IsSomething(row.leg_crushing) Or IsSomethingString(row.intermeditate_connection_type) Or IsSomething(row.intermeditate_connection_spacing) Or IsSomething(row.ki_override) Or IsSomething(row.leg_diameter) Or IsSomething(row.leg_thickness) Or IsSomething(row.leg_grade) Or IsSomething(row.leg_unbraced_length) Or IsSomething(row.rein_diameter) Or IsSomething(row.rein_thickness) Or IsSomething(row.rein_grade) Or IsSomething(row.leg_length) Or IsSomething(row.rein_length) Or IsSomething(row.set_top_to_bottom) Or IsSomething(row.flange_bolt_quantity_bot) Or IsSomething(row.flange_bolt_circle_bot) Or IsSomething(row.flange_bolt_orientation_bot) Or IsSomething(row.flange_bolt_quantity_top) Or IsSomething(row.flange_bolt_circle_top) Or IsSomething(row.flange_bolt_orientation_top) Or IsSomethingString(row.threaded_rod_size_bot) Or IsSomethingString(row.threaded_rod_mat_bot) Or IsSomething(row.threaded_rod_quantity_bot) Or IsSomething(row.threaded_rod_unbraced_length_bot) Or IsSomethingString(row.threaded_rod_size_top) Or IsSomethingString(row.threaded_rod_mat_top) Or IsSomething(row.threaded_rod_quantity_top) Or IsSomething(row.threaded_rod_unbraced_length_top) Or IsSomething(row.stiffener_height_bot) Or IsSomething(row.stiffener_length_bot) Or IsSomething(row.stiffener_fillet_bot) Or IsSomething(row.stiffener_exx_bot) Or IsSomething(row.flange_thickness_bot) Or IsSomething(row.stiffener_height_top) Or IsSomething(row.stiffener_length_top) Or IsSomething(row.stiffener_fillet_top) Or IsSomething(row.stiffener_exx_top) Or IsSomething(row.flange_thickness_top) Or IsSomethingString(row.structure_ind) Or IsSomethingString(row.leg_reinforcement_name) Or IsSomething(row.top_elev) Or IsSomething(row.bot_elev)) Then
                        SQLUpdate = SQLUpdate.Replace("--[LEG REINFORCEMENT DETAIL INSERT]", row.SQLUpdate)
                        SQLUpdate = SQLUpdate.Replace("--BEGIN --[LEG REINFORCEMENT DETAIL UPDATE BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAIL UPDATE BEGIN]")
                        SQLUpdate = SQLUpdate.Replace("--END --[LEG REINFORCEMENT DETAIL UPDATE END]", "END --[LEG REINFORCEMENT DETAIL UPDATE END]")
                    Else
                        SQLUpdate = SQLUpdate.Replace("--[LEG REINFORCEMENT DETAIL INSERT]", row.SQLDelete)
                    End If
                Else
                    If IsSomethingString(row.create_leg) Then
                        SQLUpdate = SQLUpdate.Replace("--[LEG REINFORCEMENT DETAIL INSERT]", row.SQLInsert)
                        SQLUpdate = SQLUpdate.Replace("--BEGIN --[LEG REINFORCEMENT DETAIL INSERT BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAIL INSERT BEGIN]")
                        SQLUpdate = SQLUpdate.Replace("--END --[LEG REINFORCEMENT DETAIL INSERT END]", "END --[LEG REINFORCEMENT DETAIL INSERT END]")
                    End If
                End If
            Next
        End If

        'note: additional update commands are imbedded within objects sharing similar relationships

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

    Private _Structural_105 As Boolean?

    <DataMember()> Public Property LegReinforcementDetails As New List(Of LegReinforcementDetail)


    <Category("Leg Reinforcements"), Description(""), DisplayName("Structural 105")>
    <DataMember()> Public Property Structural_105() As Boolean?
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

    Public Sub New(ByVal dr As DataRow, ByRef strDS As DataSet, Optional ByVal Parent As EDSObject = Nothing) 'Added strDS in order to pull EDS data from subtables
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        'Following used to create dataset, regardless if source was EDS or Excel. Boolean used to identify source. EDS = True
        BuildFromDataset(dr, strDS, True, Me)

    End Sub 'Generate Leg Reinforcement from EDS

    Public Sub New(ExcelFilePath As String, Optional ByVal Parent As EDSObject = Nothing)
        Me.WorkBookPath = ExcelFilePath
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        LoadFromExcel()

    End Sub 'Generate Leg Reinforcement from Excel

    Private Sub BuildFromDataset(ByVal dr As DataRow, ByRef ds As DataSet, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing)
        'Dataset is pulled in from either EDS or Excel. True = EDS, False = Excel
        'If Parent IsNot Nothing Then Me.Absorb(Parent) 'Do not double absorb!!!

        'Not sure this is necessary, could just read the values from the structure code criteria when creating the Excel sheet (Added to Save to Excel Section)
        'Me.tia_current = Me.ParentStructure?.structureCodeCriteria?.tia_current
        'Me.rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
        'Me.seismic_design_category = Me.ParentStructure?.structureCodeCriteria?.seismic_design_category

        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.Version = DBtoStr(dr.Item("tool_version"))
        Me.bus_unit = If(EDStruefalse, DBtoStr(dr.Item("bus_unit")), Me.bus_unit) 'Not provided in Excel
        Me.work_order_seq_num = If(EDStruefalse, Me.work_order_seq_num, Me.work_order_seq_num) 'Not provided in Excel
        Me.structure_id = If(EDStruefalse, DBtoStr(dr.Item("structure_id")), Me.structure_id) 'Not provided in Excel
        Me.modified_person_id = If(EDStruefalse, DBtoNullableInt(dr.Item("modified_person_id")), Me.modified_person_id) 'Not provided in Excel
        Me.process_stage = If(EDStruefalse, DBtoStr(dr.Item("process_stage")), Me.process_stage) 'Not provided in Excel
        Me.Structural_105 = DBtoNullableBool(dr.Item("Structural_105"))

        Dim lrDetails As New LegReinforcementDetail 'Leg Reinforcement Details
        'Dim lrResult As New LegReinforcementResults2 'Leg Reinforcement Results
        'Dim lrResult1 As New LegReinforcementResults 'Leg Reinforcement Results 'original

        For Each lrrow As DataRow In ds.Tables(lrDetails.EDSObjectName).Rows
            'create a new connection based on the datarow from above
            lrDetails = New LegReinforcementDetail(lrrow, ds, EDStruefalse, Me)
            'Check if the parent id, in the case leg reinforcement id is equal to the original object id (Me)                    
            If If(EDStruefalse, lrDetails.leg_reinforcement_id = Me.ID, True) Then 'If coming from Excel, all leg reinforcement details provided will be associated to leg reinforcement. 
                'If it is equal then add the newly created leg reinforcment detail to the list of details 
                LegReinforcementDetails.Add(lrDetails)
                'Results.Add(lrDetails)
            End If
        Next

    End Sub

#End Region

#Region "Load From Excel"
    Public Overrides Sub LoadFromExcel()
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet

        For Each item As EXCELDTParameter In ExcelDTParams
            'Get additional tables from excel file 
            Try
                excelDS.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(Me.WorkBookPath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
            Catch ex As Exception
                Debug.Print(String.Format("Failed to create datatable for: {0}, {1}, {2}", IO.Path.GetFileName(Me.WorkBookPath), item.xlsSheet, item.xlsRange))
            End Try
        Next

        If excelDS.Tables.Contains("Leg Reinforcements") Then
            Dim dr = excelDS.Tables("Leg Reinforcements").Rows(0)

            'Following used to create dataset, regardless if source was EDS or Excel. Boolean used to identify source. Excel = False
            BuildFromDataset(dr, excelDS, False, Me)

        End If
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

            'App ID & Revision #
            .Worksheets("IMPORT").Range("Order_Import").Value = MyOrder()

            'Tower Type - Defaulting to Self-Support if can't determine
            'Tool is set up where importing tnx file determines tower type. Pulling this from the site code criteria might not be necessary since importing geometry is required.
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.structure_type) Then
            If Me.ParentStructure?.structureCodeCriteria?.structure_type = "SELF SUPPORT" Then
                structure_type = "Self Support"
            ElseIf Me.ParentStructure?.structureCodeCriteria?.structure_type = "GUYED" Then
                structure_type = "Guyed"
            Else
                structure_type = "Self Support"
            End If
            .Worksheets("TNX File").Range("A2").Value = "TowerType=" & CType(structure_type, String)
            'End If

            'TIA Revision- Defaulting to Rev. H if not available. Currently pulled in from TNX file
            .Worksheets("TNX File").Range("A1").Value = "SteelCode=" & MyTIA(True)

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
                'identify first row to copy data into Excel Sheet
                'TNX File Data (Row 1 & 2 are used to save tia revision and tower type, respectively.)
                Dim tnxdataRow As Integer = 3 'TNX File Data tab

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

                    'Storing following in order to populate arbitrary shape info after worksheet is opened. 
                    If Not IsNothing(row.top_elev) Then
                        .Worksheets("TNX File").Range("A" & tnxdataRow).Value = "TowerRec=" & CType(row.local_id, Integer)
                        .Worksheets("TNX File").Range("A" & tnxdataRow + 1).Value = "TowerHeight=" & CType(row.top_elev, Double)
                        tnxdataRow += 2
                        'Else
                        '    .Worksheets("TNX File").Range("A" & tnxdataRow).ClearContents
                    End If
                    If Not IsNothing(row.bot_elev) Then
                        .Worksheets("TNX File").Range("A" & tnxdataRow).Value = "TowerSectionLength=" & CType(row.top_elev, Double) - CType(row.bot_elev, Double)
                        tnxdataRow += 1
                        'Else
                        '    .Worksheets("TNX File").Range("A" & tnxdataRow).ClearContents
                    End If
                Next

                'Adding following unique line to TNX File to address previous issues and help determine final line of text on Tab. 
                .Worksheets("TNX File").Range("A" & tnxdataRow).Value = "[EndOverwrite]"

            End If

        End With

    End Sub

#End Region

#Region "Save to EDS"

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Version.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Structural_105.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Structural_105")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tool_version = " & Me.Version.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit = " & Me.bus_unit.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id = " & Me.structure_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
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
        Equals = If(Me.Version.CheckChange(otherToCompare.Version, changes, categoryName, "Tool Version"), Equals, False)
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

    Public Overrides Sub Clear()
        Me.LegReinforcementDetails.Clear()
        Me.Results.Clear()
    End Sub

End Class

<DataContractAttribute()>
Partial Public Class LegReinforcementDetail
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Leg Reinforcement Details"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "tnx.memb_leg_reinforcement_details"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableDepth As Integer
        Get
            Return 1 'base 0 so this is first sub level
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        SQLInsert = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Detail_INSERT
        SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT DETAIL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT DETAIL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        'Leg Reinforcement Results
        'For Each row As LegReinforcementResults2 In LegReinforcementResults2
        '    SQLInsert = SQLInsert.Replace("--BEGIN --[LEG REINFORCEMENT DETAILS RESULTS INSERT BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAILS RESULTS INSERT BEGIN]")
        '    SQLInsert = SQLInsert.Replace("--END --[LEG REINFORCEMENT DETAILS RESULTS INSERT END]", "END --[LEG REINFORCEMENT DETAILS RESULTS INSERT END]")
        '    'SQLInsert = SQLInsert.Replace("--[LEG REINFORCEMENT DETAILS RESULTS INSERT]", row.SQLInsert)
        '    SQLInsert = SQLInsert.Replace("--[LEG REINFORCEMENT DETAILS RESULTS INSERT]", Me.Results.EDSResultQuery)
        'Next

        If Me.Results.Count > 0 Then
            SQLInsert = SQLInsert.Replace("--BEGIN --[LEG REINFORCEMENT DETAILS RESULTS INSERT BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAILS RESULTS INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[LEG REINFORCEMENT DETAILS RESULTS INSERT END]", "END --[LEG REINFORCEMENT DETAILS RESULTS INSERT END]")
            SQLInsert = SQLInsert.Replace("--[LEG REINFORCEMENT DETAILS RESULTS INSERT]", Me.Results.EDSResultQuery)
        End If

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        SQLUpdate = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Detail_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        'Leg Reinforcement Results
        'For Each row As LegReinforcementResults2 In LegReinforcementResults2
        '    SQLUpdate = SQLUpdate.Replace("--BEGIN --[LEG REINFORCEMENT DETAILS RESULTS INSERT BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAILS RESULTS INSERT BEGIN]")
        '    SQLUpdate = SQLUpdate.Replace("--END --[LEG REINFORCEMENT DETAILS RESULTS INSERT END]", "END --[LEG REINFORCEMENT DETAILS RESULTS INSERT END]")
        '    'SQLUpdate = SQLUpdate.Replace("--[LEG REINFORCEMENT DETAILS RESULTS INSERT]", row.SQLInsert)
        'Next

        If Me.Results.Count > 0 Then
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[LEG REINFORCEMENT DETAILS RESULTS INSERT BEGIN]", "BEGIN --[LEG REINFORCEMENT DETAILS RESULTS INSERT BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[LEG REINFORCEMENT DETAILS RESULTS INSERT END]", "END --[LEG REINFORCEMENT DETAILS RESULTS INSERT END]")
            SQLUpdate = SQLUpdate.Replace("--[LEG REINFORCEMENT DETAILS RESULTS INSERT]", Me.Results.EDSResultQuery)
        End If

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
    Private _structure_ind As String 'Indicator used to differentiate between upper and lower structure. Currently not implemented. 
    Private _reinforcement_type As String
    Private _leg_reinforcement_name As String
    Private _local_id As Integer?
    Private _top_elev As Double?
    Private _bot_elev As Double?
    Private _create_leg As String 'excel only field

    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Id")>
    <DataMember()> Public Property leg_reinforcement_id() As Integer?
        Get
            Return Me._leg_reinforcement_id
        End Get
        Set
            Me._leg_reinforcement_id = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Load Time Mod Option")>
    <DataMember()> Public Property leg_load_time_mod_option() As Boolean?
        Get
            Return Me._leg_load_time_mod_option
        End Get
        Set
            Me._leg_load_time_mod_option = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("End Connection Type")>
    <DataMember()> Public Property end_connection_type() As String
        Get
            Return Me._end_connection_type
        End Get
        Set
            Me._end_connection_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Crushing")>
    <DataMember()> Public Property leg_crushing() As Boolean?
        Get
            Return Me._leg_crushing
        End Get
        Set
            Me._leg_crushing = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Applied Load Type")>
    <DataMember()> Public Property applied_load_type() As String
        Get
            Return Me._applied_load_type
        End Get
        Set
            Me._applied_load_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Slenderness Ratio Type")>
    <DataMember()> Public Property slenderness_ratio_type() As String
        Get
            Return Me._slenderness_ratio_type
        End Get
        Set
            Me._slenderness_ratio_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Intermeditate Connection Type")>
    <DataMember()> Public Property intermeditate_connection_type() As String
        Get
            Return Me._intermeditate_connection_type
        End Get
        Set
            Me._intermeditate_connection_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Intermeditate Connection Spacing")>
    <DataMember()> Public Property intermeditate_connection_spacing() As Double?
        Get
            Return Me._intermeditate_connection_spacing
        End Get
        Set
            Me._intermeditate_connection_spacing = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Ki Override")>
    <DataMember()> Public Property ki_override() As Double?
        Get
            Return Me._ki_override
        End Get
        Set
            Me._ki_override = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Diameter")>
    <DataMember()> Public Property leg_diameter() As Double?
        Get
            Return Me._leg_diameter
        End Get
        Set
            Me._leg_diameter = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Thickness")>
    <DataMember()> Public Property leg_thickness() As Double?
        Get
            Return Me._leg_thickness
        End Get
        Set
            Me._leg_thickness = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Grade")>
    <DataMember()> Public Property leg_grade() As Double?
        Get
            Return Me._leg_grade
        End Get
        Set
            Me._leg_grade = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Unbraced Length")>
    <DataMember()> Public Property leg_unbraced_length() As Double?
        Get
            Return Me._leg_unbraced_length
        End Get
        Set
            Me._leg_unbraced_length = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Rein Diameter")>
    <DataMember()> Public Property rein_diameter() As Double?
        Get
            Return Me._rein_diameter
        End Get
        Set
            Me._rein_diameter = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Rein Thickness")>
    <DataMember()> Public Property rein_thickness() As Double?
        Get
            Return Me._rein_thickness
        End Get
        Set
            Me._rein_thickness = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Rein Grade")>
    <DataMember()> Public Property rein_grade() As Double?
        Get
            Return Me._rein_grade
        End Get
        Set
            Me._rein_grade = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Print Bolt On Connections")>
    <DataMember()> Public Property print_bolt_on_connections() As Boolean?
        Get
            Return Me._print_bolt_on_connections
        End Get
        Set
            Me._print_bolt_on_connections = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Length")>
    <DataMember()> Public Property leg_length() As Double?
        Get
            Return Me._leg_length
        End Get
        Set
            Me._leg_length = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Rein Length")>
    <DataMember()> Public Property rein_length() As Double?
        Get
            Return Me._rein_length
        End Get
        Set
            Me._rein_length = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Set Top To Bottom")>
    <DataMember()> Public Property set_top_to_bottom() As Boolean?
        Get
            Return Me._set_top_to_bottom
        End Get
        Set
            Me._set_top_to_bottom = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Quantity Bot")>
    <DataMember()> Public Property flange_bolt_quantity_bot() As Integer?
        Get
            Return Me._flange_bolt_quantity_bot
        End Get
        Set
            Me._flange_bolt_quantity_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Circle Bot")>
    <DataMember()> Public Property flange_bolt_circle_bot() As Double?
        Get
            Return Me._flange_bolt_circle_bot
        End Get
        Set
            Me._flange_bolt_circle_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Orientation Bot")>
    <DataMember()> Public Property flange_bolt_orientation_bot() As Integer?
        Get
            Return Me._flange_bolt_orientation_bot
        End Get
        Set
            Me._flange_bolt_orientation_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Quantity Top")>
    <DataMember()> Public Property flange_bolt_quantity_top() As Integer?
        Get
            Return Me._flange_bolt_quantity_top
        End Get
        Set
            Me._flange_bolt_quantity_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Circle Top")>
    <DataMember()> Public Property flange_bolt_circle_top() As Double?
        Get
            Return Me._flange_bolt_circle_top
        End Get
        Set
            Me._flange_bolt_circle_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Bolt Orientation Top")>
    <DataMember()> Public Property flange_bolt_orientation_top() As Integer?
        Get
            Return Me._flange_bolt_orientation_top
        End Get
        Set
            Me._flange_bolt_orientation_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Size Bot")>
    <DataMember()> Public Property threaded_rod_size_bot() As String
        Get
            Return Me._threaded_rod_size_bot
        End Get
        Set
            Me._threaded_rod_size_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Mat Bot")>
    <DataMember()> Public Property threaded_rod_mat_bot() As String
        Get
            Return Me._threaded_rod_mat_bot
        End Get
        Set
            Me._threaded_rod_mat_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Quantity Bot")>
    <DataMember()> Public Property threaded_rod_quantity_bot() As Integer?
        Get
            Return Me._threaded_rod_quantity_bot
        End Get
        Set
            Me._threaded_rod_quantity_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Unbraced Length Bot")>
    <DataMember()> Public Property threaded_rod_unbraced_length_bot() As Double?
        Get
            Return Me._threaded_rod_unbraced_length_bot
        End Get
        Set
            Me._threaded_rod_unbraced_length_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Size Top")>
    <DataMember()> Public Property threaded_rod_size_top() As String
        Get
            Return Me._threaded_rod_size_top
        End Get
        Set
            Me._threaded_rod_size_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Mat Top")>
    <DataMember()> Public Property threaded_rod_mat_top() As String
        Get
            Return Me._threaded_rod_mat_top
        End Get
        Set
            Me._threaded_rod_mat_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Quantity Top")>
    <DataMember()> Public Property threaded_rod_quantity_top() As Integer?
        Get
            Return Me._threaded_rod_quantity_top
        End Get
        Set
            Me._threaded_rod_quantity_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Threaded Rod Unbraced Length Top")>
    <DataMember()> Public Property threaded_rod_unbraced_length_top() As Double?
        Get
            Return Me._threaded_rod_unbraced_length_top
        End Get
        Set
            Me._threaded_rod_unbraced_length_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Height Bot")>
    <DataMember()> Public Property stiffener_height_bot() As Double?
        Get
            Return Me._stiffener_height_bot
        End Get
        Set
            Me._stiffener_height_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Length Bot")>
    <DataMember()> Public Property stiffener_length_bot() As Double?
        Get
            Return Me._stiffener_length_bot
        End Get
        Set
            Me._stiffener_length_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Fillet Bot")>
    <DataMember()> Public Property stiffener_fillet_bot() As Integer?
        Get
            Return Me._stiffener_fillet_bot
        End Get
        Set
            Me._stiffener_fillet_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Exx Bot")>
    <DataMember()> Public Property stiffener_exx_bot() As Double?
        Get
            Return Me._stiffener_exx_bot
        End Get
        Set
            Me._stiffener_exx_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Thickness Bot")>
    <DataMember()> Public Property flange_thickness_bot() As Double?
        Get
            Return Me._flange_thickness_bot
        End Get
        Set
            Me._flange_thickness_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Height Top")>
    <DataMember()> Public Property stiffener_height_top() As Double?
        Get
            Return Me._stiffener_height_top
        End Get
        Set
            Me._stiffener_height_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Length Top")>
    <DataMember()> Public Property stiffener_length_top() As Double?
        Get
            Return Me._stiffener_length_top
        End Get
        Set
            Me._stiffener_length_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Fillet Top")>
    <DataMember()> Public Property stiffener_fillet_top() As Integer?
        Get
            Return Me._stiffener_fillet_top
        End Get
        Set
            Me._stiffener_fillet_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Stiffener Exx Top")>
    <DataMember()> Public Property stiffener_exx_top() As Double?
        Get
            Return Me._stiffener_exx_top
        End Get
        Set
            Me._stiffener_exx_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Flange Thickness Top")>
    <DataMember()> Public Property flange_thickness_top() As Double?
        Get
            Return Me._flange_thickness_top
        End Get
        Set
            Me._flange_thickness_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Structure Ind")>
    <DataMember()> Public Property structure_ind() As String
        Get
            Return Me._structure_ind
        End Get
        Set
            Me._structure_ind = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Reinforcement Type")>
    <DataMember()> Public Property reinforcement_type() As String
        Get
            Return Me._reinforcement_type
        End Get
        Set
            Me._reinforcement_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Name")>
    <DataMember()> Public Property leg_reinforcement_name() As String
        Get
            Return Me._leg_reinforcement_name
        End Get
        Set
            Me._leg_reinforcement_name = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Local Id")>
    <DataMember()> Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Top Elev")>
    <DataMember()> Public Property top_elev() As Double?
        Get
            Return Me._top_elev
        End Get
        Set
            Me._top_elev = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Bot Elev")>
    <DataMember()> Public Property bot_elev() As Double?
        Get
            Return Me._bot_elev
        End Get
        Set
            Me._bot_elev = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Create Leg")>
    <DataMember()> Public Property create_leg() As String
        Get
            Return Me._create_leg
        End Get
        Set
            Me._create_leg = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    'Public Sub New(ByVal row As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal prow As DataRow, ByRef strDS As DataSet)
    Public Sub New(ByVal row As DataRow, ByRef ds As DataSet, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal prow As DataRow, ByRef strDS As DataSet)
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

        If EDStruefalse = False Then 'Only Pull in when referencing Excel
            Try
                Me.create_leg = DBtoStr(dr.Item("create_leg"))
            Catch ex As Exception
                Me.create_leg = Nothing
            End Try
        End If

        'only store when populating from an Excel sheet
        If EDStruefalse = False Then
            For Each resrow As DataRow In ds.Tables("Leg Reinforcement Results").Rows
                Dim local_id As Integer = CType(resrow.Item("local_id"), Integer)
                If local_id = Me.local_id Then
                    Me.Results.Add(New EDSResult(resrow, Me))
                End If
            Next

            'Needed to add this in because it was simpler that editing the SQL database. 
            'The results table is not named appropriately based on how the EDSresult is constructed. 
            For Each res In Me.Results
                res.ForeignKeyName = "memb_leg_reinforcement_details_id"
            Next
        End If


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

'Partial Public Class LegReinforcementResults
'    Inherits EDSObjectWithQueries
'    'Inherits EDSResult

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Leg Reinforcement Results"
'    Public Overrides ReadOnly Property EDSTableName As String = "tnx.memb_leg_reinforcement_results"
'    'Public Overrides ReadOnly Property EDSTableDepth As Integer = 2

'    Public Overrides Function SQLInsert() As String

'        SQLInsert = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Details_Results_INSERT
'        SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT DETAILS RESULT VALUES]", Me.SQLInsertValues)
'        SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT DETAILS RESULT FIELDS]", Me.SQLInsertFields)
'        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'        Return SQLInsert

'    End Function

'#End Region

'#Region "Define"
'    Private _leg_reinforcement_details_id As Integer?
'    Private _local_id As Integer?
'    'Private _work_order_seq_num As Double? 'not provided in Excel
'    Private _rating As Double?
'    Private _result_lkup As String
'    'Private _modified_person_id As Integer? 'not provided in Excel
'    'Private _process_stage As String 'not provided in Excel
'    'Private _modified_date As DateTime? 'not provided in Excel

'    <Category("Leg Reinforcement Results"), Description(""), DisplayName("Leg Reinforcement Details Id")>
'     <DataMember()> Public Property leg_reinforcement_details_id() As Integer?
'        Get
'            Return Me._leg_reinforcement_details_id
'        End Get
'        Set
'            Me._leg_reinforcement_details_id = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement Results"), Description(""), DisplayName("Local Id")>
'     <DataMember()> Public Property local_id() As Integer?
'        Get
'            Return Me._local_id
'        End Get
'        Set
'            Me._local_id = Value
'        End Set
'    End Property
'    '<Category("Leg Reinforcement Results"), Description(""), DisplayName("Work Order Seq Num")>
'    ' <DataMember()> Public Property work_order_seq_num() As Double?
'    '    Get
'    '        Return Me._work_order_seq_num
'    '    End Get
'    '    Set
'    '        Me._work_order_seq_num = Value
'    '    End Set
'    'End Property
'    <Category("Leg Reinforcement Results"), Description(""), DisplayName("Rating")>
'     <DataMember()> Public Property rating() As Double?
'        Get
'            Return Me._rating
'        End Get
'        Set
'            Me._rating = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement Results"), Description(""), DisplayName("Result Lkup")>
'     <DataMember()> Public Property result_lkup() As String
'        Get
'            Return Me._result_lkup
'        End Get
'        Set
'            Me._result_lkup = Value
'        End Set
'    End Property
'    '<Category("Leg Reinforcement Results"), Description(""), DisplayName("Modified Person Id")>
'    ' <DataMember()> Public Property modified_person_id() As Integer?
'    '    Get
'    '        Return Me._modified_person_id
'    '    End Get
'    '    Set
'    '        Me._modified_person_id = Value
'    '    End Set
'    'End Property
'    '<Category("Leg Reinforcement Results"), Description(""), DisplayName("Process Stage")>
'    ' <DataMember()> Public Property process_stage() As String
'    '    Get
'    '        Return Me._process_stage
'    '    End Get
'    '    Set
'    '        Me._process_stage = Value
'    '    End Set
'    'End Property
'    '<Category("Leg Reinforcement Results"), Description(""), DisplayName("Modified Date")>
'    ' <DataMember()> Public Property modified_date() As DateTime?
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

'    Public Sub New(ByVal lrrrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = lrrrow

'        Me.leg_reinforcement_details_id = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_id = DBtoNullableInt(dr.Item("local_id"))
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

'        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
'        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_reinforcement_details_id.ToString.FormatDBValue)
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

'        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_reinforcement_details_id")
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
'        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_reinforcement_details_id = " & Me.leg_reinforcement_details_id.ToString.FormatDBValue)
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
'        Dim otherToCompare As LegReinforcementResults = TryCast(other, LegReinforcementResults)
'        If otherToCompare Is Nothing Then Return False


'    End Function
'#End Region

'End Class

'Below used for testing will delete before SAPI release
'Partial Public Class LegReinforcementResults2
'    'Inherits EDSObjectWithQueries
'    Inherits EDSResult

'#Region "Inheritted"
'    Public Overrides ReadOnly Property EDSObjectName As String = "Leg Reinforcement Results"
'    'Public Overrides ReadOnly Property EDSTableName As String = "tnx.memb_leg_reinforcement_details_results"
'    'Public Overrides ReadOnly Property EDSTableDepth As Integer = 2

'    'Public Overrides Function SQLInsert() As String

'    '    SQLInsert = CCI_Engineering_Templates.My.Resources.Leg_Reinforcement_Details_Results_INSERT
'    '    SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT DETAILS RESULT VALUES]", Me.SQLInsertValues)
'    '    SQLInsert = SQLInsert.Replace("[LEG REINFORCEMENT DETAILS RESULT FIELDS]", Me.SQLInsertFields)
'    '    SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

'    '    Return SQLInsert

'    'End Function

'#End Region

'#Region "Define"
'    Private _leg_reinforcement_details_id As Integer?
'    Private _local_id As Integer?
'    'Private _work_order_seq_num As Double? 'not provided in Excel
'    Private _rating As Double?
'    Private _result_lkup As String
'    'Private _modified_person_id As Integer? 'not provided in Excel
'    'Private _process_stage As String 'not provided in Excel
'    'Private _modified_date As DateTime? 'not provided in Excel
'    Private _EDSTableName As String

'    <Category("Leg Reinforcement Results"), Description(""), DisplayName("Leg Reinforcement Details Id")>
'     <DataMember()> Public Property leg_reinforcement_details_id() As Integer?
'        Get
'            Return Me._leg_reinforcement_details_id
'        End Get
'        Set
'            Me._leg_reinforcement_details_id = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement Results"), Description(""), DisplayName("Local Id")>
'     <DataMember()> Public Property local_id() As Integer?
'        Get
'            Return Me._local_id
'        End Get
'        Set
'            Me._local_id = Value
'        End Set
'    End Property
'    '<Category("Leg Reinforcement Results"), Description(""), DisplayName("Work Order Seq Num")>
'    ' <DataMember()> Public Property work_order_seq_num() As Double?
'    '    Get
'    '        Return Me._work_order_seq_num
'    '    End Get
'    '    Set
'    '        Me._work_order_seq_num = Value
'    '    End Set
'    'End Property
'    <Category("Leg Reinforcement Results"), Description(""), DisplayName("Rating")>
'     <DataMember()> Public Property rating() As Double?
'        Get
'            Return Me._rating
'        End Get
'        Set
'            Me._rating = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement Results"), Description(""), DisplayName("Result Lkup")>
'     <DataMember()> Public Property result_lkup() As String
'        Get
'            Return Me._result_lkup
'        End Get
'        Set
'            Me._result_lkup = Value
'        End Set
'    End Property
'    '<Category("Leg Reinforcement Results"), Description(""), DisplayName("Modified Person Id")>
'    ' <DataMember()> Public Property modified_person_id() As Integer?
'    '    Get
'    '        Return Me._modified_person_id
'    '    End Get
'    '    Set
'    '        Me._modified_person_id = Value
'    '    End Set
'    'End Property
'    '<Category("Leg Reinforcement Results"), Description(""), DisplayName("Process Stage")>
'    ' <DataMember()> Public Property process_stage() As String
'    '    Get
'    '        Return Me._process_stage
'    '    End Get
'    '    Set
'    '        Me._process_stage = Value
'    '    End Set
'    'End Property
'    '<Category("Leg Reinforcement Results"), Description(""), DisplayName("Modified Date")>
'    ' <DataMember()> Public Property modified_date() As DateTime?
'    '    Get
'    '        Return Me._modified_date
'    '    End Get
'    '    Set
'    '        Me._modified_date = Value
'    '    End Set
'    'End Property
'    <Category("Leg Reinforcement Results"), Description(""), DisplayName("EDS Table Name")>
'     <DataMember()> Public Property EDSTableName() As String
'        Get
'            Return Me._EDSTableName
'        End Get
'        Set
'            Me._EDSTableName = Value
'        End Set
'    End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Modified Person Id")>
'    ' <DataMember()> Public Property modified_person_id() As Integer?
'    '    Get
'    '        Return Me._modified_person_id
'    '    End Get
'    '    Set
'    '        Me._modified_person_id = Value
'    '    End Set
'    'End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Process Stage")>
'    ' <DataMember()> Public Property process_stage() As String
'    '    Get
'    '        Return Me._process_stage
'    '    End Get
'    '    Set
'    '        Me._process_stage = Value
'    '    End Set
'    'End Property
'    '<Category("Bolt Results"), Description(""), DisplayName("Modified Date")>
'    ' <DataMember()> Public Property modified_date() As DateTime?
'    '    Get
'    '        Return Me._modified_date
'    '    End Get
'    '    Set
'    '        Me._modified_date = Value
'    '    End Set
'    'End Property

'#End Region

'#End Region

'    Sub New()
'        'Leave method empty
'    End Sub

'    Public Sub New(ByVal lrrrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)

'        Dim dr = lrrrow

'        Me.leg_reinforcement_details_id = DBtoNullableInt(dr.Item("ID"))
'        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
'            Me.local_id = DBtoNullableInt(dr.Item("local_id"))
'        End If
'        'Me.work_order_seq_num = DBtoNullableDbl(dr.Item("work_order_seq_num"))
'        Me.rating = DBtoNullableDbl(dr.Item("rating"))
'        Me.result_lkup = DBtoStr(dr.Item("result_lkup"))
'        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
'        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
'        'Me.modified_date = DBtoStr(dr.Item("modified_date"))
'        Me.EDSTableName = "tnx.memb_leg_reinforcement_details_results"

'    End Sub

'#End Region

'#Region "Save to EDS"
'    'Public Overrides Function SQLInsertValues() As String
'    '    SQLInsertValues = ""

'    '    SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
'    '    'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_reinforcement_details_id.ToString.FormatDBValue)
'    '    'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
'    '    SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
'    '    SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
'    '    'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
'    '    'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
'    '    'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)


'    '    Return SQLInsertValues
'    'End Function

'    'Public Overrides Function SQLInsertFields() As String
'    '    SQLInsertFields = ""

'    '    SQLInsertFields = SQLInsertFields.AddtoDBString("leg_reinforcement_details_id")
'    '    'SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
'    '    SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
'    '    SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
'    '    'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
'    '    'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
'    '    'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")


'    '    Return SQLInsertFields
'    'End Function

'    'Public Overrides Function SQLUpdateFieldsandValues() As String
'    '    SQLUpdateFieldsandValues = ""
'    '    SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_reinforcement_details_id = " & Me.leg_reinforcement_details_id.ToString.FormatDBValue)
'    '    'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
'    '    SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
'    '    SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
'    '    'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
'    '    'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
'    '    'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

'    '    Return SQLUpdateFieldsandValues
'    'End Function
'#End Region

'#Region "Equals"
'    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
'        Equals = True
'        If changes Is Nothing Then changes = New List(Of AnalysisChange)
'        Dim categoryName As String = Me.EDSObjectFullName

'        'Makes sure you are comparing to the same object type
'        'Customize this to the object type
'        Dim otherToCompare As LegReinforcementResults = TryCast(other, LegReinforcementResults)
'        If otherToCompare Is Nothing Then Return False


'    End Function
'#End Region

'End Class