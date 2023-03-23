Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
'Imports Microsoft.Office.Interop

Partial Public Class Pole
    Inherits EDSExcelObject

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole General"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.pole"
    Public Overrides ReadOnly Property templatePath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "CCIpole.xlsm")

    Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("CCIpole General EXCEL", "A2:K3", "General (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Pole Sections EXCEL", "A2:W20", "Unreinf Pole (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Pole Reinf Sections EXCEL", "A2:W102", "Reinf Pole (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Reinf Groups EXCEL", "A2:J50", "Reinf Groups (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Reinf Details EXCEL", "A2:H200", "Reinf ID (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Int Groups EXCEL", "A2:H50", "Interference Groups (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Int Details EXCEL", "A2:H200", "Interference ID (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Pole Reinf Results EXCEL", "A2:I1000", "Reinf Results (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Matl Property Details EXCEL", "A2:G20", "Materials (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Bolt Property Details EXCEL", "A2:S20", "Bolts (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Reinf Property Details EXCEL", "A2:EB50", "Reinforcements (SAPI)")}
            '***Add additional table references here****
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        If _Insert = "" Then
            '_Insert = QueryBuilderFromFile(queryPath & "CCIpole\1 General (INSERT).sql")
            _Insert = CCI_Engineering_Templates.My.Resources.CCIpole_General_INSERT
        End If
        SQLInsert = _Insert

        'General Pole Data
        SQLInsert = SQLInsert.Replace("[GENERAL POLE VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[GENERAL POLE FIELDS]", Me.SQLInsertFields)


        'Unreinforced Sections
        For Each row As PoleSection In unreinf_sections
            SQLInsert = SQLInsert.Replace("--[UNREINF SECTION SUBQUERY]", row.SQLInsert)
        Next

        'Reinf Sections
        For Each row As PoleReinfSection In reinf_sections
            SQLInsert = SQLInsert.Replace("--[REINF SECTION SUBQUERY]", row.SQLInsert)
        Next

        'Reinforcement Groups
        For Each row As PoleReinfGroup In reinf_groups
            SQLInsert = SQLInsert.Replace("--[REINF GROUP SUBQUERY]", row.SQLInsert)
        Next
        'Details are done in query for Groups

        'Interference Groups
        For Each row As PoleIntGroup In int_groups
            SQLInsert = SQLInsert.Replace("--[INT GROUP SUBQUERY]", row.SQLInsert)
        Next
        'Details are done in query for Groups

        'Results
        For Each row As PoleReinfResults In reinf_section_results
            SQLInsert = SQLInsert.Replace("--[RESULT SUBQUERY]", row.SQLInsert)
        Next

        'Reinf/Bolt/Matl DBs are done within sections/group queries - No its not - Yes it is.

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        'This section not only needs to call update commands but also needs to call insert and delete commands since subtables may involve adding or deleting records
        If _Update = "" Then
            '_Update = QueryBuilderFromFile(queryPath & "CCIpole\1 General (UPDATE).sql")
            _Update = CCI_Engineering_Templates.My.Resources.CCIpole_General_UPDATE
        End If
        SQLUpdate = _Update

        'General Pole Data
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.pole_id.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)


        'Unreinforced Sections
        For Each row As PoleSection In unreinf_sections
            If IsSomething(row.section_id) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                If IsSomething(row.local_section_id) Then
                    SQLUpdate = SQLUpdate.Replace("--[UNREINF SECTION SUBQUERY]", row.SQLUpdate)
                Else
                    SQLUpdate = SQLUpdate.Replace("--[UNREINF SECTION SUBQUERY]", row.SQLDelete)
                End If
            Else
                SQLUpdate = SQLUpdate.Replace("--[UNREINF SECTION SUBQUERY]", row.SQLInsert)
            End If
        Next

        'Reinf Sections
        For Each row As PoleReinfSection In reinf_sections
            If IsSomething(row.section_id) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                If IsSomething(row.local_section_id) Then
                    SQLUpdate = SQLUpdate.Replace("--[REINF SECTION SUBQUERY]", row.SQLUpdate)
                Else
                    SQLUpdate = SQLUpdate.Replace("--[REINF SECTION SUBQUERY]", row.SQLDelete)
                End If
            Else
                SQLUpdate = SQLUpdate.Replace("--[REINF SECTION SUBQUERY]", row.SQLInsert)
            End If
        Next

        'Reinforcement Groups
        For Each row As PoleReinfGroup In reinf_groups
            If IsSomething(row.group_id) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                If IsSomething(row.local_group_id) Then
                    SQLUpdate = SQLUpdate.Replace("--[REINF GROUP SUBQUERY]", row.SQLUpdate)
                Else
                    SQLUpdate = SQLUpdate.Replace("--[REINF GROUP SUBQUERY]", row.SQLDelete)
                End If
            Else
                SQLUpdate = SQLUpdate.Replace("--[REINF GROUP SUBQUERY]", row.SQLInsert)
            End If
        Next
        'Details are done in query for Groups

        'Interference Groups
        For Each row As PoleIntGroup In int_groups
            If IsSomething(row.group_id) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                If IsSomething(row.local_group_id) Then
                    SQLUpdate = SQLUpdate.Replace("--[INT GROUP SUBQUERY]", row.SQLUpdate)
                Else
                    SQLUpdate = SQLUpdate.Replace("--[INT GROUP SUBQUERY]", row.SQLDelete)
                End If
            Else
                SQLUpdate = SQLUpdate.Replace("--[INT GROUP SUBQUERY]", row.SQLInsert)
            End If
        Next
        'Details are done in query for Groups

        'Results
        SQLUpdate = SQLUpdate.Replace("--[RESULT SUBQUERY]", "DELETE FROM pole.reinforcement_results WHERE work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue & vbNewLine & "--[RESULT SUBQUERY]")
        For Each row As PoleReinfResults In reinf_section_results
            SQLUpdate = SQLUpdate.Replace("--[RESULT SUBQUERY]", row.SQLInsert)
            'If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
            '    If IsSomething(row.local_section_id) Or IsSomething(row.local_group_id) Then
            '        SQLUpdate = SQLUpdate.Replace("--[RESULTS SUBQUERY]", row.SQLUpdate)
            '    Else
            '        SQLUpdate = SQLUpdate.Replace("--[RESULTS SUBQUERY]", row.SQLDelete)
            '    End If
            'Else
            '    SQLUpdate = SQLUpdate.Replace("--[RESULTS SUBQUERY]", row.SQLInsert)
            'End If
            'SQLUpdate = SQLUpdate.Replace("--[RESULTS UPDATE]", Me.Results.EDSResultQuery) 'Chall using EDSResultQuery. Not sure if necessary in CCIpole...we will see - MRR
        Next

        'Reinf/Bolt/Matl DBs are done within sections/group queries

    End Function

    Public Overrides Function SQLDelete() As String
        If _Delete = "" Then
            '_Delete = QueryBuilderFromFile(queryPath & "CCIpole\1 General (DELETE).sql")
            _Delete = CCI_Engineering_Templates.My.Resources.CCIpole_General_DELETE
        End If
        SQLDelete = _Delete

        'General Pole Data
        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\General (DELETE).sql") 'CWH - previously ran into issues when _Delete = String which is why this code was used. 
        SQLDelete = SQLDelete.Replace("[ID]", Me.pole_id.ToString.FormatDBValue)


        'Unreinforced Sections
        For Each row As PoleSection In unreinf_sections
            SQLDelete = SQLDelete.Replace("--[UNREINF SECTION SUBQUERY]", row.SQLDelete)
        Next

        'Reinf Sections
        For Each row As PoleReinfSection In reinf_sections
            SQLDelete = SQLDelete.Replace("--[REINF SECTION SUBQUERY]", row.SQLDelete)
        Next

        'Reinforcement Groups
        For Each row As PoleReinfGroup In reinf_groups
            SQLDelete = SQLDelete.Replace("--[REINF GROUP SUBQUERY]", row.SQLDelete)
        Next
        'Details are done in query for Groups

        'Interference Groups
        For Each row As PoleIntGroup In int_groups
            SQLDelete = SQLDelete.Replace("--[INT GROUP SUBQUERY]", row.SQLDelete)
        Next
        'Details are done in query for Groups

        'Results
        SQLDelete = SQLDelete.Replace("--[RESULTS SUBQUERY]", "DELETE FROM pole.reinforcement_results WHERE pole_id = " & Me.pole_id)
        'For Each row As PoleReinfResults In reinf_section_results
        '    SQLDelete = SQLDelete.Replace("[RESULTS SUBQUERY]", row.SQLDelete)
        'Next

        'Reinf/Bolt/Matl DBs should not be deleted from the DB except for manual cases

    End Function
#End Region

#Region "Define"
    'Private _bus_unit As String
    'Private _structure_id As String
    Private _pole_id As Integer?
    Private _upper_structure_type As String
    Private _analysis_deg As Double?
    Private _geom_increment_length As Double?
    Private _tool_version As String
    Private _check_connections As Boolean?
    Private _hole_deformation As Boolean?
    Private _ineff_mod_check As Boolean?
    Private _modified As Boolean?
    'Private _modified_person_id As Integer?
    'Private _process_stage As String

    Public Property unreinf_sections As New List(Of PoleSection)
    Public Property reinf_sections As New List(Of PoleReinfSection)
    Public Property reinf_groups As New List(Of PoleReinfGroup)
    'Public Property reinf_ids As New List(Of PoleReinfDetail)
    Public Property int_groups As New List(Of PoleIntGroup)
    'Public Property int_ids As New List(Of PoleIntDetail)
    Public Property reinf_section_results As New List(Of PoleReinfResults)
    Public Property matls As New List(Of PoleMatlProp)
    Public Property bolts As New List(Of PoleBoltProp)
    Public Property reinfs As New List(Of PoleReinfProp)



    '<Category("Pole"), Description(""), DisplayName("Bus Unit")>
    'Public Property bus_unit() As String
    '    Get
    '        Return Me._bus_unit
    '    End Get
    '    Set
    '        Me._bus_unit = Value
    '    End Set
    'End Property
    '<Category("Pole"), Description(""), DisplayName("Structure Id")>
    'Public Property structure_id() As String
    '    Get
    '        Return Me._structure_id
    '    End Get
    '    Set
    '        Me._structure_id = Value
    '    End Set
    'End Property
    <Category("Pole"), Description(""), DisplayName("Pole Id")>
    Public Property pole_id() As Integer?
        Get
            Return Me._pole_id
        End Get
        Set
            Me._pole_id = Value
        End Set
    End Property
    <Category("Pole"), Description(""), DisplayName("Upper Structure Type")>
    Public Property upper_structure_type() As String
        Get
            Return Me._upper_structure_type
        End Get
        Set
            Me._upper_structure_type = Value
        End Set
    End Property
    <Category("Pole"), Description(""), DisplayName("Analysis Deg")>
    Public Property analysis_deg() As Double?
        Get
            Return Me._analysis_deg
        End Get
        Set
            Me._analysis_deg = Value
        End Set
    End Property
    <Category("Pole"), Description(""), DisplayName("Geom Increment Length")>
    Public Property geom_increment_length() As Double?
        Get
            Return Me._geom_increment_length
        End Get
        Set
            Me._geom_increment_length = Value
        End Set
    End Property
    <Category("Pole"), Description(""), DisplayName("Tool Version")>
    Public Property tool_version() As String
        Get
            Return Me._tool_version
        End Get
        Set
            Me._tool_version = Value
        End Set
    End Property
    <Category("Pole"), Description(""), DisplayName("Check Connections")>
    Public Property check_connections() As Boolean?
        Get
            Return Me._check_connections
        End Get
        Set
            Me._check_connections = Value
        End Set
    End Property
    <Category("Pole"), Description(""), DisplayName("Hole Deformation")>
    Public Property hole_deformation() As Boolean?
        Get
            Return Me._hole_deformation
        End Get
        Set
            Me._hole_deformation = Value
        End Set
    End Property
    <Category("Pole"), Description(""), DisplayName("Ineff Mod Check")>
    Public Property ineff_mod_check() As Boolean?
        Get
            Return Me._ineff_mod_check
        End Get
        Set
            Me._ineff_mod_check = Value
        End Set
    End Property
    <Category("Pole"), Description(""), DisplayName("Modified")>
    Public Property modified() As Boolean?
        Get
            Return Me._modified
        End Get
        Set
            Me._modified = Value
        End Set
    End Property
    '<Category("Pole"), Description(""), DisplayName("Modified Person Id")>
    'Public Property modified_person_id() As Integer?
    '    Get
    '        Return Me._modified_person_id
    '    End Get
    '    Set
    '        Me._modified_person_id = Value
    '    End Set
    'End Property
    '<Category("Pole"), Description(""), DisplayName("Process Stage")>
    'Public Property process_stage() As String
    '    Get
    '        Return Me._process_stage
    '    End Get
    '    Set
    '        Me._process_stage = Value
    '    End Set
    'End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal dr As DataRow, ByRef strDS As DataSet, Optional ByVal Parent As EDSObject = Nothing) 'Added strDS in order to pull EDS data from subtables
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        'Get values from structure code criteria
        'Not sure this is necessary, could just read the values from the structure code criteria when creating the Excel sheet (Added to Save to Excel Section)
        'Me.tia_current = Me.ParentStructure?.structureCodeCriteria?.tia_current
        'Me.rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
        'If Me.tia_current <> "TIA-222-H" Then Me.load_z = True
        'Me.work_order_seq_num = Me.ParentStructure?.structureCodeCriteria?.work_order_seq_num

        ''''''Customize for each foundation type'''''
        Me.bus_unit = DBtoStr(dr.Item("bus_unit"))
        Me.structure_id = DBtoStr(dr.Item("structure_id"))
        Me.pole_id = DBtoNullableInt(dr.Item("ID"))
        Me.ID = Me.pole_id 'Needed for EDSListQueryBuilder Insert/Update/Delete Determination - MRR
        Me.upper_structure_type = DBtoStr(dr.Item("upper_structure_type"))
        Me.analysis_deg = DBtoNullableDbl(dr.Item("analysis_deg"))
        Me.geom_increment_length = DBtoNullableDbl(dr.Item("geom_increment_length"))
        Me.tool_version = DBtoStr(dr.Item("tool_version"))
        Me.check_connections = DBtoNullableBool(dr.Item("check_connections"))
        Me.hole_deformation = DBtoNullableBool(dr.Item("hole_deformation"))
        Me.ineff_mod_check = DBtoNullableBool(dr.Item("ineff_mod_check"))
        Me.modified = DBtoNullableBool(dr.Item("modified"))
        Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
        Me.process_stage = DBtoStr(dr.Item("process_stage"))

        'Sub Tables (as defined within Structure.VB LoadFromEDS
        For Each dbrow As DataRow In strDS.Tables("Pole Custom Matls").Rows
            'Dim PropMatlRefID As Integer? = CType(row.Item("pole_id"), Integer)
            'If PropMatlRefID = Me.ID Then
            Me.matls.Add(New PoleMatlProp(dbrow, Me))
            'End If
        Next 'Add Custom Matl Properties to CCIpole Object

        For Each dbrow As DataRow In strDS.Tables("Pole Custom Bolts").Rows
            Me.bolts.Add(New PoleBoltProp(dbrow, Me))
        Next 'Add Custom Bolt Properties to CCIpole Object

        For Each dbrow As DataRow In strDS.Tables("Pole Custom Reinfs").Rows
            Me.reinfs.Add(New PoleReinfProp(dbrow, Me))
        Next 'Add Custom Reinf Properties to CCIpole Object


        For Each row As DataRow In strDS.Tables("Pole Unreinforced Sections").Rows
            Dim PoleSectionRefID As Integer? = CType(row.Item("pole_id"), Integer)
            If PoleSectionRefID = Me.pole_id Then
                Me.unreinf_sections.Add(New PoleSection(row, Me)) 'Chall has 'Add(New PoleSection(row, srtDS))' - apparently keeping the datasource tied in allows the UI to show relationships correctly - Not sure if necessary yet - MRR

                ''Add any non-default Matls from DB into Object (only pulling IDs if they were used)
                'Dim SectionMatlID As Integer? = CType(row.Item("matl_id"), Integer)
                'For Each dbrow As DataRow In strDS.Tables("Pole Custom Matls").Rows
                '    Dim MatlID As Integer? = CType(dbrow.Item("ID"), Integer)
                '    Dim MatlDefault As Boolean = CType(dbrow.Item("ind_default"), Boolean)
                '    If SectionMatlID = MatlID And MatlDefault = False Then
                '        Me.matls.Add(New PoleMatlProp(dbrow, Me))
                '    End If
                'Next

            End If
        Next 'Add Unreinf Sections to CCIpole Object

        For Each row As DataRow In strDS.Tables("Pole Reinforced Sections").Rows
            Dim PoleReinfSectionRefID As Integer? = CType(row.Item("pole_id"), Integer)
            If PoleReinfSectionRefID = Me.pole_id Then
                Me.reinf_sections.Add(New PoleReinfSection(row, Me))
            End If
        Next 'Add Reinf Sections to CCIpole Object

        For Each GroupRow As DataRow In strDS.Tables("Pole Reinforcement Groups").Rows
            Dim PoleReinfGroupRefID As Integer? = CType(GroupRow.Item("pole_id"), Integer)
            Dim PoleReinfGroupID As Integer? = CType(GroupRow.Item("ID"), Integer)

            If PoleReinfGroupRefID = Me.pole_id Then
                Dim NewReinfGroup As PoleReinfGroup
                NewReinfGroup = New PoleReinfGroup(GroupRow, Me)

                For Each DetailRow As DataRow In strDS.Tables("Pole Reinforcement Details").Rows
                    Dim PoleReinfDetailRefID As Integer? = CType(DetailRow.Item("group_id"), Integer)
                    If PoleReinfDetailRefID = PoleReinfGroupID Then
                        NewReinfGroup.reinf_ids.Add(New PoleReinfDetail(DetailRow, NewReinfGroup))
                    End If
                Next 'Add Reinf Details to Group Object

                Me.reinf_groups.Add(NewReinfGroup)
            End If
        Next 'Add Reinf Groups to CCIpole Object

        For Each GroupRow As DataRow In strDS.Tables("Pole Interference Groups").Rows
            Dim PoleIntGroupRefID As Integer? = CType(GroupRow.Item("pole_id"), Integer)
            Dim PoleIntGroupID As Integer? = CType(GroupRow.Item("ID"), Integer)

            If PoleIntGroupRefID = Me.pole_id Then
                Dim NewIntGroup As PoleIntGroup
                NewIntGroup = New PoleIntGroup(GroupRow, Me)

                For Each DetailRow As DataRow In strDS.Tables("Pole Interference Details").Rows
                    Dim PoleIntDetailsRefID As Integer? = CType(DetailRow.Item("group_id"), Integer)
                    If PoleIntDetailsRefID = PoleIntGroupID Then
                        NewIntGroup.int_ids.Add(New PoleIntDetail(DetailRow, NewIntGroup))
                    End If
                Next 'Add Interference Details to Group Object

                Me.int_groups.Add(NewIntGroup)
            End If
        Next 'Add Interference Groups to CCIpole Object


        'For Each row As DataRow In strDS.Tables("Pole Results").Rows
        '    Dim PoleReinfResultsRefID As Integer? = CType(row.Item("pole_id"), Integer)
        '    If PoleReinfResultsRefID = Me.ID Then
        '        Me.reinf_section_results.Add(New PoleReinfResults(row, Me))
        '    End If
        'Next 'Add Reinf Section Results to CCIpole Object


    End Sub 'Generate Pole from EDS

    Public Sub New(ExcelFilePath As String, Optional ByVal Parent As EDSObject = Nothing)
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

        If excelDS.Tables.Contains("CCIpole General EXCEL") Then
            Dim dr = excelDS.Tables("CCIpole General EXCEL").Rows(0)
            'Need to dimension DataRow from GenStructure/TNX and anywhere else inputs may come from as well - MRR

            'Me.bus_unit = DBtoStr(dr.Item("bus_unit"))
            'Me.structure_id = DBtoStr(dr.Item("structure_id"))
            Me.pole_id = DBtoNullableInt(dr.Item("ID"))
            Me.ID = Me.pole_id 'Needed for EDSListQueryBuilder Insert/Update/Delete Determination - MRR
            Me.upper_structure_type = DBtoStr(dr.Item("upper_structure_type"))
            Me.analysis_deg = DBtoNullableDbl(dr.Item("analysis_deg"))
            Me.geom_increment_length = DBtoNullableDbl(dr.Item("geom_increment_length"))
            Me.tool_version = DBtoStr(dr.Item("tool_version"))
            Me.check_connections = DBtoNullableBool(dr.Item("check_connections"))
            Me.hole_deformation = DBtoNullableBool(dr.Item("hole_deformation"))
            Me.ineff_mod_check = DBtoNullableBool(dr.Item("ineff_mod_check"))
            Me.modified = DBtoNullableBool(dr.Item("modified"))
            'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
            'Me.process_stage = DBtoStr(dr.Item("process_stage"))
        End If


        If excelDS.Tables.Contains("CCIpole Matl Property Details EXCEL") Then
            For Each row As DataRow In excelDS.Tables("CCIpole Matl Property Details EXCEL").Rows
                Me.matls.Add(New PoleMatlProp(row, Me))
            Next
        End If

        If excelDS.Tables.Contains("CCIpole Bolt Property Details EXCEL") Then
            For Each row As DataRow In excelDS.Tables("CCIpole Bolt Property Details EXCEL").Rows
                Me.bolts.Add(New PoleBoltProp(row, Me))
            Next
        End If

        If excelDS.Tables.Contains("CCIpole Reinf Property Details EXCEL") Then
            For Each row As DataRow In excelDS.Tables("CCIpole Reinf Property Details EXCEL").Rows
                Me.reinfs.Add(New PoleReinfProp(row, Me))
            Next
        End If


        If excelDS.Tables.Contains("CCIpole Pole Sections EXCEL") Then
            For Each row As DataRow In excelDS.Tables("CCIpole Pole Sections EXCEL").Rows
                Me.unreinf_sections.Add(New PoleSection(row, Me))
            Next
        End If

        If excelDS.Tables.Contains("CCIpole Pole Reinf Sections EXCEL") Then
            For Each row As DataRow In excelDS.Tables("CCIpole Pole Reinf Sections EXCEL").Rows
                Me.reinf_sections.Add(New PoleReinfSection(row, Me))
            Next
        End If

        If excelDS.Tables.Contains("CCIpole Reinf Groups EXCEL") Then
            For Each GroupRow As DataRow In excelDS.Tables("CCIpole Reinf Groups EXCEL").Rows
                Dim NewReinfGroup As PoleReinfGroup
                NewReinfGroup = New PoleReinfGroup(GroupRow, Me)

                Dim ReinfGroupID, DetailsGroupID As Integer?

                Try
                    If Not IsDBNull(CType(GroupRow.Item("local_group_id"), Integer)) Then
                        ReinfGroupID = CType(GroupRow.Item("local_group_id"), Integer)
                    Else
                        ReinfGroupID = Nothing
                    End If
                Catch
                    ReinfGroupID = Nothing
                End Try 'Group local_group_id

                For Each DetailRow As DataRow In excelDS.Tables("CCIpole Reinf Details EXCEL").Rows

                    Try
                        If Not IsDBNull(CType(DetailRow.Item("local_group_id"), Integer)) Then
                            DetailsGroupID = CType(DetailRow.Item("local_group_id"), Integer)
                        Else
                            DetailsGroupID = Nothing
                        End If
                    Catch
                        DetailsGroupID = Nothing
                    End Try 'Details local_group_id

                    If ReinfGroupID = DetailsGroupID Then
                        NewReinfGroup.reinf_ids.Add(New PoleReinfDetail(DetailRow, NewReinfGroup))
                    End If

                Next

                Me.reinf_groups.Add(NewReinfGroup)

            Next

        End If


        If excelDS.Tables.Contains("CCIpole Int Groups EXCEL") Then
            For Each GroupRow As DataRow In excelDS.Tables("CCIpole Int Groups EXCEL").Rows
                Dim NewIntGroup As PoleIntGroup
                NewIntGroup = New PoleIntGroup(GroupRow, Me)

                Dim IntGroupID, IntDetailsGroupID As Integer?

                Try
                    If Not IsDBNull(CType(GroupRow.Item("local_group_id"), Integer)) Then
                        IntGroupID = CType(GroupRow.Item("local_group_id"), Integer)
                    Else
                        IntGroupID = Nothing
                    End If
                Catch
                    IntGroupID = Nothing
                End Try 'Group local_group_id

                For Each DetailRow As DataRow In excelDS.Tables("CCIpole Int Details EXCEL").Rows

                    Try
                        If Not IsDBNull(CType(DetailRow.Item("local_group_id"), Integer)) Then
                            IntDetailsGroupID = CType(DetailRow.Item("local_group_id"), Integer)
                        Else
                            IntDetailsGroupID = Nothing
                        End If
                    Catch
                        IntDetailsGroupID = Nothing
                    End Try 'Details local_group_id

                    If IntGroupID = IntDetailsGroupID Then
                        NewIntGroup.int_ids.Add(New PoleIntDetail(DetailRow, NewIntGroup))
                    End If

                Next

                Me.int_groups.Add(NewIntGroup)

            Next

        End If

        If excelDS.Tables.Contains("CCIpole Pole Reinf Results EXCEL") Then
            For Each row As DataRow In excelDS.Tables("CCIpole Pole Reinf Results EXCEL").Rows
                Me.reinf_section_results.Add(New PoleReinfResults(row, Me))
            Next
        End If


    End Sub 'Generate Pole from Excel
#End Region

#Region "Save to Excel"
    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        '''''Customize for each excel tool'''''

        With wb

            'CCIpole General
            Dim pole_tia_current As String
            Dim pole_rev_h_section_15_5 As Boolean?

            If Not IsNothing(Me.bus_unit) Then
                .Worksheets("General (SAPI)").Range("A3").Value = CType(Me.bus_unit, Integer)
            Else
                .Worksheets("General (SAPI)").Range("A3").ClearContents
            End If
            If Not IsNothing(Me.structure_id) Then
                .Worksheets("General (SAPI)").Range("B3").Value = CType(Me.structure_id, String)
            End If
            If Not IsNothing(Me.pole_id) Then
                .Worksheets("General (SAPI)").Range("C3").Value = CType(Me.pole_id, Integer)
            Else
                .Worksheets("General (SAPI)").Range("C3").ClearContents
            End If
            If Not IsNothing(Me.upper_structure_type) Then
                .Worksheets("General (SAPI)").Range("D3").Value = CType(Me.upper_structure_type, String)
            End If
            If Not IsNothing(Me.analysis_deg) Then
                .Worksheets("General (SAPI)").Range("E3").Value = CType(Me.analysis_deg, Double)
            Else
                .Worksheets("General (SAPI)").Range("E3").ClearContents
            End If
            If Not IsNothing(Me.geom_increment_length) Then
                .Worksheets("General (SAPI)").Range("F3").Value = CType(Me.geom_increment_length, Double)
            Else
                .Worksheets("General (SAPI)").Range("F3").ClearContents
            End If
            If Not IsNothing(Me.tool_version) Then
                .Worksheets("General (SAPI)").Range("G3").Value = CType(Me.tool_version, String)
            End If
            If Not IsNothing(Me.check_connections) Then
                .Worksheets("General (SAPI)").Range("H3").Value = CType(Me.check_connections, Boolean)
            End If
            If Not IsNothing(Me.hole_deformation) Then
                .Worksheets("General (SAPI)").Range("I3").Value = CType(Me.hole_deformation, Boolean)
            End If
            If Not IsNothing(Me.ineff_mod_check) Then
                .Worksheets("General (SAPI)").Range("J3").Value = CType(Me.ineff_mod_check, Boolean)
            End If
            If Not IsNothing(Me.modified) Then
                .Worksheets("General (SAPI)").Range("K3").Value = CType(Me.modified, Boolean)
            End If
            'If Not IsNothing(Me.modified_person_id) Then
            '    .Worksheets("General (SAPI)").Range("L3").Value = CType(Me.modified_person_id, Integer)
            'Else
            '    .Worksheets("General (SAPI)").Range("L3").ClearContents
            'End If
            'If Not IsNothing(Me.process_stage) Then
            '    .Worksheets("General (SAPI)").Range("M3").Value = CType(Me.process_stage, String)
            'End If

            'Site Code Critera
            'TIA Revision- Defaulting to Rev. H if not available. 
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.tia_current) Then
                If Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-F" Then
                    pole_tia_current = "F"
                ElseIf Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-G" Then
                    pole_tia_current = "G"
                Else
                    pole_tia_current = "H"
                End If
            Else
                pole_tia_current = "H"
            End If
            .Worksheets("General (SAPI)").Range("P3").Value = CType(pole_tia_current, String)

            'Load Z Normalization
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.load_z_norm) Then
            '    rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.load_z_norm
            '    .Worksheets("General (SAPI)").Range("Q3").Value = CType(load_z_norm, Boolean)
            'End If
            'H Section 15.5
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5) Then
                pole_rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
                .Worksheets("General (SAPI)").Range("R3").Value = CType(pole_rev_h_section_15_5, Boolean)
            End If
            'Work Order
            If Not IsNothing(Me.ParentStructure?.work_order_seq_num) Then
                work_order_seq_num = Me.ParentStructure?.work_order_seq_num
                .Worksheets("General (SAPI)").Range("S3").Value = CType(work_order_seq_num, Integer)
            End If

            Dim row As Integer = 2
            Dim col As Integer = 0
            Dim drow, dcol As Integer

            'Unreinforced Pole Sections
            For Each ps As PoleSection In unreinf_sections
                col = 0

                If Not IsNothing(ps.section_id) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.section_id, Integer)
                col += 1
                If Not IsNothing(Me.pole_id) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(Me.pole_id, Integer)
                col += 1
                If Not IsNothing(ps.local_section_id) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.local_section_id, Integer)
                col += 1
                If Not IsNothing(ps.elev_bot) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.elev_bot, Double)
                col += 1
                If Not IsNothing(ps.elev_top) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.elev_top, Double)
                col += 1
                If Not IsNothing(ps.length_section) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.length_section, Double)
                col += 1
                If Not IsNothing(ps.length_splice) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.length_splice, Double)
                col += 1
                If Not IsNothing(ps.num_sides) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.num_sides, Integer)
                col += 1
                If Not IsNothing(ps.diam_bot) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.diam_bot, Double)
                col += 1
                If Not IsNothing(ps.diam_top) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.diam_top, Double)
                col += 1
                If Not IsNothing(ps.wall_thickness) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.wall_thickness, Double)
                col += 1
                If ps.bend_radius = -1 Then
                    .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = "Auto"
                ElseIf Not IsNothing(ps.bend_radius) Then
                    .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.bend_radius, Double)
                End If
                col += 1
                If Not IsNothing(ps.matl_id) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.matl_id, Integer)
                col += 1
                If Not IsNothing(ps.local_matl_id) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.local_matl_id, Integer)
                col += 1
                'If Not IsNothing(ps.pole_type) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.pole_type, String)
                col += 1
                If Not IsNothing(ps.section_name) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.section_name, String)
                col += 1
                'If Not IsNothing(ps.socket_length) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.socket_length, Double)
                'col += 1
                'If Not IsNothing(ps.weight_mult) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.weight_mult, Double)
                'col += 1
                'If Not IsNothing(ps.wp_mult) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.wp_mult, Double)
                'col += 1
                'If Not IsNothing(ps.af_factor) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.af_factor, Double)
                'col += 1
                'If Not IsNothing(ps.ar_factor) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.ar_factor, Double)
                'col += 1
                'If Not IsNothing(ps.round_area_ratio) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.round_area_ratio, Double)
                'col += 1
                'If Not IsNothing(ps.flat_area_ratio) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.flat_area_ratio, Double)

                row += 1
            Next

            row = 2

            'Reinforced Pole Sections - Does this need to be Imported from EDS into Excel? - MRR
            For Each prs As PoleReinfSection In reinf_sections

                col = 0

                If Not IsNothing(prs.section_id) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.section_id, Integer)
                col += 1
                If Not IsNothing(Me.pole_id) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(Me.pole_id, Integer)
                col += 1
                If Not IsNothing(prs.local_section_id) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.local_section_id, Integer)
                col += 1
                If Not IsNothing(prs.elev_bot) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.elev_bot, Double)
                col += 1
                If Not IsNothing(prs.elev_top) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.elev_top, Double)
                col += 1
                If Not IsNothing(prs.length_section) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.length_section, Double)
                col += 1
                If Not IsNothing(prs.length_splice) Then
                    .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.length_splice, Double)
                Else
                    .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = ""
                End If
                col += 1
                If Not IsNothing(prs.num_sides) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.num_sides, Integer)
                col += 1
                If Not IsNothing(prs.diam_bot) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.diam_bot, Double)
                col += 1
                If Not IsNothing(prs.diam_top) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.diam_top, Double)
                col += 1
                If Not IsNothing(prs.wall_thickness) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.wall_thickness, Double)
                col += 1
                If Not IsNothing(prs.bend_radius) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.bend_radius, Double)
                col += 1
                If Not IsNothing(prs.matl_id) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.matl_id, Integer)
                col += 1
                If Not IsNothing(prs.local_matl_id) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.local_matl_id, Integer)
                col += 1
                'If Not IsNothing(prs.pole_type) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = prs.pole_type
                col += 1
                If Not IsNothing(prs.weight_mult) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.weight_mult, Double)
                col += 1
                If Not IsNothing(prs.section_name) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = prs.section_name
                col += 1
                'If Not IsNothing(prs.socket_length) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.socket_length, Double)
                'col += 1
                'If Not IsNothing(prs.wp_mult) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.wp_mult, Double)
                'col += 1
                'If Not IsNothing(prs.af_factor) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.af_factor, Double)
                'col += 1
                'If Not IsNothing(prs.ar_factor) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.ar_factor, Double)
                'col += 1
                'If Not IsNothing(prs.round_area_ratio) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.round_area_ratio, Double)
                'col += 1
                'If Not IsNothing(prs.flat_area_ratio) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.flat_area_ratio, Double)

                row += 1

            Next

            row = 2
            drow = 2

            'Reinforcement Groups
            For Each prg As PoleReinfGroup In reinf_groups

                col = 0

                If Not IsNothing(prg.group_id) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.group_id, Integer)
                col += 1
                If Not IsNothing(Me.pole_id) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(Me.pole_id, Integer)
                col += 1
                If Not IsNothing(prg.local_group_id) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.local_group_id, Integer)
                col += 1
                If Not IsNothing(prg.elev_bot_actual) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.elev_bot_actual, Double)
                col += 1
                If Not IsNothing(prg.elev_bot_eff) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.elev_bot_eff, Double)
                col += 1
                If Not IsNothing(prg.elev_top_actual) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.elev_top_actual, Double)
                col += 1
                If Not IsNothing(prg.elev_top_eff) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.elev_top_eff, Double)
                col += 1
                If Not IsNothing(prg.reinf_id) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.reinf_id, Integer)
                col += 1
                If Not IsNothing(prg.local_reinf_id) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.local_reinf_id, Integer)
                col += 1
                If Not IsNothing(prg.qty) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.qty, Integer)


                'Individual Reinforcement Details 
                For Each prd As PoleReinfDetail In prg.reinf_ids

                    If prd.local_group_id = prg.local_group_id Then

                        dcol = 0

                        If Not IsNothing(prd.reinforcement_id) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.reinforcement_id, Integer)
                        dcol += 1
                        If Not IsNothing(prg.group_id) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prg.group_id, Integer)
                        dcol += 1
                        If Not IsNothing(prd.local_group_id) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.local_group_id, Integer)
                        dcol += 1
                        If Not IsNothing(prd.local_reinforcement_id) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.local_reinforcement_id, Integer)
                        dcol += 1
                        If Not IsNothing(prd.pole_flat) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.pole_flat, Integer)
                        dcol += 1
                        If Not IsNothing(prd.horizontal_offset) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.horizontal_offset, Double)
                        dcol += 1
                        If Not IsNothing(prd.rotation) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.rotation, Double)
                        dcol += 1
                        If Not IsNothing(prd.note) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = prd.note

                        drow += 1

                    End If

                Next

                row += 1

            Next

            row = 2
            drow = 2

            'Interference Groups
            For Each pig As PoleIntGroup In int_groups

                col = 0

                If Not IsNothing(pig.group_id) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.group_id, Integer)
                col += 1
                If Not IsNothing(Me.pole_id) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(Me.pole_id, Integer)
                col += 1
                If Not IsNothing(pig.local_group_id) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.local_group_id, Integer)
                col += 1
                If Not IsNothing(pig.elev_bot) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.elev_bot, Double)
                col += 1
                If Not IsNothing(pig.elev_top) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.elev_top, Double)
                col += 1
                If Not IsNothing(pig.width) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.width, Double)
                col += 1
                If Not IsNothing(pig.description) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = pig.description
                col += 1
                If Not IsNothing(pig.qty) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.qty, Integer)

                'Individual Interference Details

                For Each pid As PoleIntDetail In pig.int_ids

                    If pid.local_group_id = pig.local_group_id Then

                        dcol = 0

                        If Not IsNothing(pid.interference_id) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pid.interference_id, Integer)
                        dcol += 1
                        If Not IsNothing(pig.group_id) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pig.group_id, Integer)
                        dcol += 1
                        If Not IsNothing(pig.local_group_id) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pig.local_group_id, Integer)
                        dcol += 1
                        If Not IsNothing(pid.local_interference_id) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pid.local_interference_id, Integer)
                        dcol += 1
                        If Not IsNothing(pid.pole_flat) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pid.pole_flat, Integer)
                        dcol += 1
                        If Not IsNothing(pid.horizontal_offset) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pid.horizontal_offset, Double)
                        dcol += 1
                        If Not IsNothing(pid.rotation) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pid.rotation, Double)
                        dcol += 1
                        If Not IsNothing(pid.note) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = pid.note

                        drow += 1

                    End If

                Next

                row += 1

            Next

            row = 2

            'Reinforced Pole Section Results
            For Each prr As PoleReinfResults In reinf_section_results

                col = 0

                If Not IsNothing(prr.result_id) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.result_id, Double)
                col += 1
                If Not IsNothing(prr.work_order_seq_num) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.work_order_seq_num, Double)
                col += 1
                If Not IsNothing(prr.pole_id) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.pole_id, Double)
                col += 1
                If Not IsNothing(prr.section_id) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.section_id, Integer)
                col += 1
                If Not IsNothing(prr.local_section_id) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.local_section_id, Integer)
                col += 1
                If Not IsNothing(prr.group_id) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.group_id, Integer)
                col += 1
                If Not IsNothing(prr.local_group_id) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.local_group_id, Integer)
                col += 1
                If Not IsNothing(prr.result_lkup) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.result_lkup, Integer)
                col += 1
                If Not IsNothing(prr.rating) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.rating, Double)

                row += 1

            Next

            row = 2

            'Custom Material Properties
            For Each pm As PoleMatlProp In matls

                If pm.ind_default = False Then

                    col = 0

                    If Not IsNothing(Me.pole_id) Then .Worksheets("Materials (SAPI)").Cells(row, col).Value = CType(Me.pole_id, Integer)
                    col += 1
                    If Not IsNothing(pm.matl_id) Then .Worksheets("Materials (SAPI)").Cells(row, col).Value = CType(pm.matl_id, Integer)
                    col += 1
                    If Not IsNothing(pm.local_matl_id) Then .Worksheets("Materials (SAPI)").Cells(row, col).Value = CType(pm.local_matl_id, Integer)
                    col += 1
                    If Not IsNothing(pm.name) Then .Worksheets("Materials (SAPI)").Cells(row, col).Value = pm.name
                    col += 1
                    If Not IsNothing(pm.fy) Then .Worksheets("Materials (SAPI)").Cells(row, col).Value = CType(pm.fy, Double)
                    col += 1
                    If Not IsNothing(pm.fu) Then .Worksheets("Materials (SAPI)").Cells(row, col).Value = CType(pm.fu, Double)
                    col += 1
                    If Not IsNothing(pm.ind_default) Then .Worksheets("Materials (SAPI)").Cells(row, col).Value = CType(pm.ind_default, Boolean)

                    row += 1

                End If

            Next

            row = 2

            'Custom Bolt Properties
            For Each pb As PoleBoltProp In bolts

                If pb.ind_default = False Then

                    col = 0

                    If Not IsNothing(Me.pole_id) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(Me.pole_id, Integer)
                    col += 1
                    If Not IsNothing(pb.bolt_id) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.bolt_id, Integer)
                    col += 1
                    If Not IsNothing(pb.local_bolt_id) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.local_bolt_id, Integer)
                    col += 1
                    If Not IsNothing(pb.name) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = pb.name
                    col += 1
                    If Not IsNothing(pb.description) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = pb.description
                    col += 1
                    If Not IsNothing(pb.diam) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.diam, Double)
                    col += 1
                    If Not IsNothing(pb.area) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.area, Double)
                    col += 1
                    If Not IsNothing(pb.fu_bolt) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.fu_bolt, Double)
                    col += 1
                    If Not IsNothing(pb.sleeve_diam_out) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.sleeve_diam_out, Double)
                    col += 1
                    If Not IsNothing(pb.sleeve_diam_in) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.sleeve_diam_in, Double)
                    col += 1
                    If Not IsNothing(pb.fu_sleeve) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.fu_sleeve, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_n_sleeve_shear_revF) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.bolt_n_sleeve_shear_revF, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_x_sleeve_shear_revF) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.bolt_x_sleeve_shear_revF, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_n_sleeve_shear_revG) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.bolt_n_sleeve_shear_revG, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_x_sleeve_shear_revG) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.bolt_x_sleeve_shear_revG, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_n_sleeve_shear_revH) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.bolt_n_sleeve_shear_revH, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_x_sleeve_shear_revH) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.bolt_x_sleeve_shear_revH, Double)
                    col += 1
                    If Not IsNothing(pb.rb_applied_revH) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.rb_applied_revH, Boolean)
                    col += 1
                    If Not IsNothing(pb.ind_default) Then .Worksheets("Bolts (SAPI)").Cells(row, col).Value = CType(pb.ind_default, Boolean)

                    row += 1

                End If

            Next

            row = 2

            'Custom Reinforcement Properties
            For Each pr As PoleReinfProp In reinfs

                If pr.ind_default = False Then

                    col = 0

                    If Not IsNothing(Me.pole_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(Me.pole_id, Integer)
                    col += 1
                    If Not IsNothing(pr.reinf_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.reinf_id, Integer)
                    col += 1
                    If Not IsNothing(pr.local_reinf_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.local_reinf_id, Integer)
                    col += 1
                    If Not IsNothing(pr.name) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.name
                    col += 1
                    If Not IsNothing(pr.type) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.type
                    col += 1
                    If Not IsNothing(pr.b) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.b, Double)
                    col += 1
                    If Not IsNothing(pr.h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.h, Double)
                    col += 1
                    If Not IsNothing(pr.sr_diam) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.sr_diam, Double)
                    col += 1
                    If Not IsNothing(pr.channel_thkns_web) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.channel_thkns_web, Double)
                    col += 1
                    If Not IsNothing(pr.channel_thkns_flange) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.channel_thkns_flange, Double)
                    col += 1
                    If Not IsNothing(pr.channel_eo) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.channel_eo, Double)
                    col += 1
                    If Not IsNothing(pr.channel_J) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.channel_J, Double)
                    col += 1
                    If Not IsNothing(pr.channel_Cw) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.channel_Cw, Double)
                    col += 1
                    If Not IsNothing(pr.area_gross) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.area_gross, Double)
                    col += 1
                    If Not IsNothing(pr.centroid) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.centroid, Double)
                    col += 1
                    If Not IsNothing(pr.istension) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.istension, Boolean)
                    col += 1
                    'MATL ID
                    If Not IsNothing(pr.matl_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.matl_id, Integer)
                    col += 1
                    If Not IsNothing(pr.local_matl_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.local_matl_id, Integer)
                    col += 1
                    If Not IsNothing(pr.Ix) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.Ix, Double)
                    col += 1
                    If Not IsNothing(pr.Iy) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.Iy, Double)
                    col += 1
                    If Not IsNothing(pr.Lu) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.Lu, Double)
                    col += 1
                    If Not IsNothing(pr.Kx) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.Kx, Double)
                    col += 1
                    If Not IsNothing(pr.Ky) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.Ky, Double)
                    col += 1
                    If Not IsNothing(pr.bolt_hole_size) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_hole_size, Double)
                    col += 1
                    If Not IsNothing(pr.area_net) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.area_net, Double)
                    col += 1
                    If Not IsNothing(pr.shear_lag) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.shear_lag, Double)
                    col += 1
                    If Not IsNothing(pr.connection_type_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.connection_type_bot
                    col += 1
                    If Not IsNothing(pr.connection_cap_revF_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.connection_cap_revF_bot, Double)
                    col += 1
                    If Not IsNothing(pr.connection_cap_revG_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.connection_cap_revG_bot, Double)
                    col += 1
                    If Not IsNothing(pr.connection_cap_revH_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.connection_cap_revH_bot, Double)
                    col += 1
                    'BOT BOLT ID
                    If Not IsNothing(pr.bolt_id_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_id_bot, Integer)
                    col += 1
                    If Not IsNothing(pr.local_bolt_id_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.local_bolt_id_bot, Integer)
                    col += 1
                    If Not IsNothing(pr.bolt_N_or_X_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.bolt_N_or_X_bot
                    col += 1
                    If Not IsNothing(pr.bolt_num_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_num_bot, Integer)
                    col += 1
                    If Not IsNothing(pr.bolt_spacing_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_spacing_bot, Double)
                    col += 1
                    If Not IsNothing(pr.bolt_edge_dist_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_edge_dist_bot, Double)
                    col += 1
                    If Not IsNothing(pr.FlangeOrBP_connected_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.FlangeOrBP_connected_bot, Boolean)
                    col += 1
                    If Not IsNothing(pr.weld_grade_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_grade_bot, Double)
                    col += 1
                    If Not IsNothing(pr.weld_trans_type_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.weld_trans_type_bot
                    col += 1
                    If Not IsNothing(pr.weld_trans_length_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_trans_length_bot, Double)
                    col += 1
                    If Not IsNothing(pr.weld_groove_depth_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_groove_depth_bot, Double)
                    col += 1
                    If Not IsNothing(pr.weld_groove_angle_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_groove_angle_bot, Integer)
                    col += 1
                    If Not IsNothing(pr.weld_trans_fillet_size_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_trans_fillet_size_bot, Double)
                    col += 1
                    If Not IsNothing(pr.weld_trans_eff_throat_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_trans_eff_throat_bot, Double)
                    col += 1
                    If Not IsNothing(pr.weld_long_type_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.weld_long_type_bot
                    col += 1
                    If Not IsNothing(pr.weld_long_length_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_long_length_bot, Double)
                    col += 1
                    If Not IsNothing(pr.weld_long_fillet_size_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_long_fillet_size_bot, Double)
                    col += 1
                    If Not IsNothing(pr.weld_long_eff_throat_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_long_eff_throat_bot, Double)
                    col += 1
                    If Not IsNothing(pr.top_bot_connections_symmetrical) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.top_bot_connections_symmetrical, Boolean)
                    col += 1
                    If Not IsNothing(pr.connection_type_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.connection_type_top
                    col += 1
                    If Not IsNothing(pr.connection_cap_revF_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.connection_cap_revF_top, Double)
                    col += 1
                    If Not IsNothing(pr.connection_cap_revG_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.connection_cap_revG_top, Double)
                    col += 1
                    If Not IsNothing(pr.connection_cap_revH_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.connection_cap_revH_top, Double)
                    col += 1
                    'TOP BOLT ID
                    If Not IsNothing(pr.bolt_id_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_id_top, Integer)
                    col += 1
                    If Not IsNothing(pr.local_bolt_id_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.local_bolt_id_top, Integer)
                    col += 1
                    If Not IsNothing(pr.bolt_N_or_X_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.bolt_N_or_X_top
                    col += 1
                    If Not IsNothing(pr.bolt_num_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_num_top, Integer)
                    col += 1
                    If Not IsNothing(pr.bolt_spacing_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_spacing_top, Double)
                    col += 1
                    If Not IsNothing(pr.bolt_edge_dist_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_edge_dist_top, Double)
                    col += 1
                    If Not IsNothing(pr.FlangeOrBP_connected_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.FlangeOrBP_connected_top, Boolean)
                    col += 1
                    If Not IsNothing(pr.weld_grade_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_grade_top, Double)
                    col += 1
                    If Not IsNothing(pr.weld_trans_type_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.weld_trans_type_top
                    col += 1
                    If Not IsNothing(pr.weld_trans_length_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_trans_length_top, Double)
                    col += 1
                    If Not IsNothing(pr.weld_groove_depth_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_groove_depth_top, Double)
                    col += 1
                    If Not IsNothing(pr.weld_groove_angle_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_groove_angle_top, Integer)
                    col += 1
                    If Not IsNothing(pr.weld_trans_fillet_size_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_trans_fillet_size_top, Double)
                    col += 1
                    If Not IsNothing(pr.weld_trans_eff_throat_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_trans_eff_throat_top, Double)
                    col += 1
                    If Not IsNothing(pr.weld_long_type_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.weld_long_type_top
                    col += 1
                    If Not IsNothing(pr.weld_long_length_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_long_length_top, Double)
                    col += 1
                    If Not IsNothing(pr.weld_long_fillet_size_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_long_fillet_size_top, Double)
                    col += 1
                    If Not IsNothing(pr.weld_long_eff_throat_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.weld_long_eff_throat_top, Double)
                    col += 1
                    If Not IsNothing(pr.conn_length_channel) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.conn_length_channel, Double)
                    col += 1
                    If Not IsNothing(pr.conn_length_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.conn_length_bot, Double)
                    col += 1
                    If Not IsNothing(pr.conn_length_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.conn_length_top, Double)
                    col += 1
                    If Not IsNothing(pr.cap_comp_xx_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_comp_xx_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_comp_yy_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_comp_yy_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_tens_yield_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_tens_yield_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_tens_rupture_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_tens_rupture_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_shear_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_shear_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_bolt_shear_bot_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_bolt_shear_bot_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_bolt_shear_top_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_bolt_shear_top_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_nodeform_bot_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_deform_bot_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_deform_bot_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_nodeform_top_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_nodeform_top_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_deform_top_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_deform_top_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_nodeform_bot_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_deform_bot_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_deform_bot_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_nodeform_top_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_nodeform_top_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_deform_top_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_deform_top_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_trans_bot_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_trans_bot_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_long_bot_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_long_bot_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_trans_top_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_trans_top_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_long_top_f) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_long_top_f, Double)
                    col += 1
                    If Not IsNothing(pr.cap_comp_xx_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_comp_xx_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_comp_yy_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_comp_yy_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_tens_yield_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_tens_yield_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_tens_rupture_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_tens_rupture_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_shear_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_shear_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_bolt_shear_bot_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_bolt_shear_bot_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_bolt_shear_top_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_bolt_shear_top_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_nodeform_bot_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_deform_bot_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_deform_bot_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_nodeform_top_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_nodeform_top_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_deform_top_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_deform_top_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_nodeform_bot_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_deform_bot_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_deform_bot_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_nodeform_top_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_nodeform_top_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_deform_top_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_deform_top_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_trans_bot_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_trans_bot_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_long_bot_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_long_bot_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_trans_top_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_trans_top_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_long_top_g) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_long_top_g, Double)
                    col += 1
                    If Not IsNothing(pr.cap_comp_xx_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_comp_xx_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_comp_yy_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_comp_yy_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_tens_yield_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_tens_yield_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_tens_rupture_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_tens_rupture_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_shear_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_shear_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_bolt_shear_bot_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_bolt_shear_bot_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_bolt_shear_top_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_bolt_shear_top_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_nodeform_bot_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_deform_bot_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_deform_bot_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_nodeform_top_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_nodeform_top_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltshaft_bearing_deform_top_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltshaft_bearing_deform_top_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_nodeform_bot_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_deform_bot_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_deform_bot_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_nodeform_top_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_nodeform_top_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_boltreinf_bearing_deform_top_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_boltreinf_bearing_deform_top_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_trans_bot_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_trans_bot_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_long_bot_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_long_bot_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_trans_top_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_trans_top_h, Double)
                    col += 1
                    If Not IsNothing(pr.cap_weld_long_top_h) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.cap_weld_long_top_h, Double)
                    col += 1
                    If Not IsNothing(pr.ind_default) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.ind_default, Boolean)

                    row += 1

                End If

            Next


            'Worksheet Change Events
            .Worksheets("Macro References").Range("EDS_Import").Value = True
            'Call Sub EDS_Connection_Import()

        End With

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.upper_structure_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.analysis_deg.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.geom_increment_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tool_version.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.check_connections.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.hole_deformation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ineff_mod_check.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("pole_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("upper_structure_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("analysis_deg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("geom_increment_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        SQLInsertFields = SQLInsertFields.AddtoDBString("check_connections")
        SQLInsertFields = SQLInsertFields.AddtoDBString("hole_deformation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ineff_mod_check")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit = " & Me.bus_unit.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id = " & Me.structure_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_id = " & Me.pole_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("upper_structure_type = " & Me.upper_structure_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("analysis_deg = " & Me.analysis_deg.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("geom_increment_length = " & Me.geom_increment_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tool_version = " & Me.tool_version.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("check_connections = " & Me.check_connections.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("hole_deformation = " & Me.hole_deformation.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ineff_mod_check = " & Me.ineff_mod_check.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified = " & Me.modified.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)


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
        Dim otherToCompare As Pole = TryCast(other, Pole)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.bus_unit.CheckChange(otherToCompare.bus_unit, changes, categoryName, "Bus Unit"), Equals, False)
        'Equals = If(Me.structure_id.CheckChange(otherToCompare.structure_id, changes, categoryName, "Structure Id"), Equals, False)
        'Equals = If(Me.pole_id.CheckChange(otherToCompare.pole_id, changes, categoryName, "Pole Id"), Equals, False)
        Equals = If(Me.upper_structure_type.CheckChange(otherToCompare.upper_structure_type, changes, categoryName, "Upper Structure Type"), Equals, False)
        Equals = If(Me.analysis_deg.CheckChange(otherToCompare.analysis_deg, changes, categoryName, "Analysis Deg"), Equals, False)
        Equals = If(Me.geom_increment_length.CheckChange(otherToCompare.geom_increment_length, changes, categoryName, "Geom Increment Length"), Equals, False)
        Equals = If(Me.tool_version.CheckChange(otherToCompare.tool_version, changes, categoryName, "Tool Version"), Equals, False)
        Equals = If(Me.check_connections.CheckChange(otherToCompare.check_connections, changes, categoryName, "Check Connections"), Equals, False)
        Equals = If(Me.hole_deformation.CheckChange(otherToCompare.hole_deformation, changes, categoryName, "Hole Deformation"), Equals, False)
        Equals = If(Me.ineff_mod_check.CheckChange(otherToCompare.ineff_mod_check, changes, categoryName, "Ineff Mod Check"), Equals, False)
        Equals = If(Me.modified.CheckChange(otherToCompare.modified, changes, categoryName, "Modified"), Equals, False)
        'Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        'Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)

        'Unreinforced Sections
        Equals = If(Me.unreinf_sections.CheckChange(otherToCompare.unreinf_sections, changes, categoryName, "Pole Unreinforced Sections"), Equals, False)

        'Reinforced Sections
        Equals = If(Me.reinf_sections.CheckChange(otherToCompare.reinf_sections, changes, categoryName, "Pole Reinforced Sections"), Equals, False)

        'Reinforcement Groups
        Equals = If(Me.reinf_groups.CheckChange(otherToCompare.reinf_groups, changes, categoryName, "Pole Reinforcement Groups"), Equals, False)

        'Reinforcement Details (Added to Reinf Group Equals Region)
        'Equals = If(Me.reinf_ids.CheckChange(otherToCompare.reinf_ids, changes, categoryName, "Pole Reinforcement Details"), Equals, False)

        'Interference Groups
        Equals = If(Me.int_groups.CheckChange(otherToCompare.int_groups, changes, categoryName, "Pole Interference Groups"), Equals, False)

        'Interference Details (Added to Int Group Equals Region)
        'Equals = If(Me.int_ids.CheckChange(otherToCompare.int_ids, changes, categoryName, "Pole Interference Details"), Equals, False)

        'Reinforcement Results
        Equals = If(Me.reinf_section_results.CheckChange(otherToCompare.reinf_section_results, changes, categoryName, "Pole Reinforcement Results"), Equals, False)

        'Custom Materials
        'Equals = If(Me.matls.CheckChange(otherToCompare.matls, changes, categoryName, "Pole Custom Matls"), Equals, False)

        'Custom Bolts
        'Equals = If(Me.bolts.CheckChange(otherToCompare.bolts, changes, categoryName, "Pole Custom Bolts"), Equals, False)

        'Custom Reinforcements
        'Equals = If(Me.reinfs.CheckChange(otherToCompare.reinfs, changes, categoryName, "Pole Custom Reinfs"), Equals, False)


    End Function
#End Region

End Class

Partial Public Class PoleSection
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Unreinforced Sections"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.sections"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\2 Unreinf Section (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_Unreinf_Section_INSERT

        If IsSomething(Me.matl_id) And Me.matl_id <> 0 Then
            SQLInsert = SQLInsert.Replace("[MATL ID]", Me.matl_id.ToString.FormatDBValue)
        Else
            SQLInsert = SQLInsert.Replace("[MATL ID]", "NULL")

            Dim ParentPole As Pole = TryCast(Me.Parent, Pole)
            If IsSomething(ParentPole) Then
                For Each dbrow As PoleMatlProp In ParentPole.matls
                    If Me.local_matl_id = dbrow.local_matl_id Then 'And Me.local_matl_id > 17 Then 'Matching, Non-Standard Materials
                        SQLInsert = SQLInsert.Replace("--[MATL DB SUBQUERY]", dbrow.SQLInsert)
                        'SQLInsert = SQLInsert.Replace("[MATL DB FIELDS AND VALUES]", dbrow.SQLUpdateFieldsandValues)
                        'SQLInsert = SQLInsert.Replace("[MATL DB FIELDS]", dbrow.SQLInsertFields)
                        'SQLInsert = SQLInsert.Replace("[MATL DB VALUES]", dbrow.SQLInsertValues)
                    End If
                Next
            End If

        End If

        SQLInsert = SQLInsert.Replace("[UNREINF SECTION VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[UNREINF SECTION FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\2 Unreinf Section (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIpole_Unreinf_Section_UPDATE

        If IsSomething(Me.matl_id) And Me.matl_id <> 0 Then
            SQLUpdate = SQLUpdate.Replace("[MATL ID]", Me.matl_id.ToString.FormatDBValue)
        Else
            SQLUpdate = SQLUpdate.Replace("[MATL ID]", "NULL")

            Dim ParentPole As Pole = TryCast(Me.Parent, Pole)
            If IsSomething(ParentPole) Then
                For Each dbrow As PoleMatlProp In ParentPole.matls
                    If Me.local_matl_id = dbrow.local_matl_id Then
                        SQLUpdate = SQLUpdate.Replace("--[MATL DB SUBQUERY]", dbrow.SQLInsert)
                    End If
                Next
            End If

        End If

        SQLUpdate = SQLUpdate.Replace("[ID]", Me.section_id.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\2 Unreinf Section (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIpole_Unreinf_Section_DELETE

        SQLDelete = SQLDelete.Replace("[ID]", Me.section_id.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete
    End Function
#End Region

#Region "Define"
    Private _section_id As Integer?
    Private _pole_id As Integer?
    Private _local_section_id As Integer?
    Private _elev_bot As Double?
    Private _elev_top As Double?
    Private _length_section As Double?
    Private _length_splice As Double?
    Private _num_sides As Integer?
    Private _diam_bot As Double?
    Private _diam_top As Double?
    Private _wall_thickness As Double?
    Private _bend_radius As Double?
    Private _matl_id As Integer?
    Private _local_matl_id As Integer?
    Private _pole_type As String
    Private _section_name As String
    Private _socket_length As Double?
    Private _weight_mult As Double?
    Private _wp_mult As Double?
    Private _af_factor As Double?
    Private _ar_factor As Double?
    Private _round_area_ratio As Double?
    Private _flat_area_ratio As Double?
    'Public Property matls As New List(Of PoleMatlProp) 'referencing me.parent.matls for query builder instead of having matls be a subproperty of the object

    <Category("CCIpole Sections"), Description(""), DisplayName("Section Id")>
    Public Property section_id() As Integer?
        Get
            Return Me._section_id
        End Get
        Set
            Me._section_id = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Pole Id")>
    Public Property pole_id() As Integer?
        Get
            Return Me._pole_id
        End Get
        Set
            Me._pole_id = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Local Section Id")>
    Public Property local_section_id() As Integer?
        Get
            Return Me._local_section_id
        End Get
        Set
            Me._local_section_id = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Elev Bot")>
    Public Property elev_bot() As Double?
        Get
            Return Me._elev_bot
        End Get
        Set
            Me._elev_bot = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Elev Top")>
    Public Property elev_top() As Double?
        Get
            Return Me._elev_top
        End Get
        Set
            Me._elev_top = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Length Section")>
    Public Property length_section() As Double?
        Get
            Return Me._length_section
        End Get
        Set
            Me._length_section = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Length Splice")>
    Public Property length_splice() As Double?
        Get
            Return Me._length_splice
        End Get
        Set
            Me._length_splice = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Num Sides")>
    Public Property num_sides() As Integer?
        Get
            Return Me._num_sides
        End Get
        Set
            Me._num_sides = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Diam Bot")>
    Public Property diam_bot() As Double?
        Get
            Return Me._diam_bot
        End Get
        Set
            Me._diam_bot = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Diam Top")>
    Public Property diam_top() As Double?
        Get
            Return Me._diam_top
        End Get
        Set
            Me._diam_top = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Wall Thickness")>
    Public Property wall_thickness() As Double?
        Get
            Return Me._wall_thickness
        End Get
        Set
            Me._wall_thickness = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Bend Radius")>
    Public Property bend_radius() As Double?
        Get
            Return Me._bend_radius
        End Get
        Set
            Me._bend_radius = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Matl Id")>
    Public Property matl_id() As Integer?
        Get
            Return Me._matl_id
        End Get
        Set
            Me._matl_id = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Local Matl Id")>
    Public Property local_matl_id() As Integer?
        Get
            Return Me._local_matl_id
        End Get
        Set
            Me._local_matl_id = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Pole Type")>
    Public Property pole_type() As String
        Get
            Return Me._pole_type
        End Get
        Set
            Me._pole_type = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Section Name")>
    Public Property section_name() As String
        Get
            Return Me._section_name
        End Get
        Set
            Me._section_name = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Socket Length")>
    Public Property socket_length() As Double?
        Get
            Return Me._socket_length
        End Get
        Set
            Me._socket_length = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Weight Mult")>
    Public Property weight_mult() As Double?
        Get
            Return Me._weight_mult
        End Get
        Set
            Me._weight_mult = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Wp Mult")>
    Public Property wp_mult() As Double?
        Get
            Return Me._wp_mult
        End Get
        Set
            Me._wp_mult = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Af Factor")>
    Public Property af_factor() As Double?
        Get
            Return Me._af_factor
        End Get
        Set
            Me._af_factor = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Ar Factor")>
    Public Property ar_factor() As Double?
        Get
            Return Me._ar_factor
        End Get
        Set
            Me._ar_factor = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Round Area Ratio")>
    Public Property round_area_ratio() As Double?
        Get
            Return Me._round_area_ratio
        End Get
        Set
            Me._round_area_ratio = Value
        End Set
    End Property
    <Category("CCIpole Sections"), Description(""), DisplayName("Flat Area Ratio")>
    Public Property flat_area_ratio() As Double?
        Get
            Return Me._flat_area_ratio
        End Get
        Set
            Me._flat_area_ratio = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByVal Parent As EDSObject = Nothing) 'ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        'Dim excelDS As New DataSet
        Dim dr = Row

        Me.section_id = DBtoNullableInt(dr.Item("ID"))
        Me.pole_id = DBtoNullableInt(dr.Item("pole_id"))
        Me.local_section_id = DBtoNullableInt(dr.Item("local_section_id"))
        Me.elev_bot = DBtoNullableDbl(dr.Item("elev_bot"))
        Me.elev_top = DBtoNullableDbl(dr.Item("elev_top"))
        Me.length_section = DBtoNullableDbl(dr.Item("length_section"))
        Me.length_splice = DBtoNullableDbl(dr.Item("length_splice"))
        Me.num_sides = DBtoNullableInt(dr.Item("num_sides"))
        Me.diam_bot = DBtoNullableDbl(dr.Item("diam_bot"))
        Me.diam_top = DBtoNullableDbl(dr.Item("diam_top"))
        Me.wall_thickness = DBtoNullableDbl(dr.Item("wall_thickness"))
        Me.bend_radius = If(DBtoStr(dr.Item("bend_radius")) = "Auto", -1, DBtoNullableDbl(dr.Item("bend_radius")))
        Me.matl_id = DBtoNullableInt(dr.Item("matl_id"))
        Me.local_matl_id = DBtoNullableInt(dr.Item("local_matl_id"))
        Me.pole_type = DBtoStr(dr.Item("pole_type"))
        Me.section_name = DBtoStr(dr.Item("section_name"))
        Me.socket_length = DBtoNullableDbl(dr.Item("socket_length"))
        Me.weight_mult = DBtoNullableDbl(dr.Item("weight_mult"))
        Me.wp_mult = DBtoNullableDbl(dr.Item("wp_mult"))
        Me.af_factor = DBtoNullableDbl(dr.Item("af_factor"))
        Me.ar_factor = DBtoNullableDbl(dr.Item("ar_factor"))
        Me.round_area_ratio = DBtoNullableDbl(dr.Item("round_area_ratio"))
        Me.flat_area_ratio = DBtoNullableDbl(dr.Item("flat_area_ratio"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.section_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") '(Me.pole_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_section_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.length_section.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.length_splice.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.num_sides.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.diam_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.diam_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.wall_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bend_radius.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel4ID") '(Me.matl_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_matl_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.section_name.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.socket_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weight_mult.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.wp_mult.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.af_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ar_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.round_area_ratio.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flat_area_ratio.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("length_section")
        SQLInsertFields = SQLInsertFields.AddtoDBString("length_splice")
        SQLInsertFields = SQLInsertFields.AddtoDBString("num_sides")
        SQLInsertFields = SQLInsertFields.AddtoDBString("diam_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("diam_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("wall_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bend_radius")
        SQLInsertFields = SQLInsertFields.AddtoDBString("matl_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_matl_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("section_name")
        SQLInsertFields = SQLInsertFields.AddtoDBString("socket_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weight_mult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("wp_mult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("af_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ar_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("round_area_ratio")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flat_area_ratio")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        Dim ParentPole As Pole = TryCast(Me.Parent, Pole)

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("section_id = " & Me.section_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_id = " & ParentPole.pole_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_section_id = " & Me.local_section_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_bot = " & Me.elev_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_top = " & Me.elev_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("length_section = " & Me.length_section.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("length_splice = " & Me.length_splice.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("num_sides = " & Me.num_sides.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("diam_bot = " & Me.diam_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("diam_top = " & Me.diam_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("wall_thickness = " & Me.wall_thickness.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bend_radius = " & Me.bend_radius.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("matl_id = @SubLevel4ID") '& Me.matl_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_matl_id = " & Me.local_matl_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_type = " & Me.pole_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("section_name = " & Me.section_name.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("socket_length = " & Me.socket_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weight_mult = " & Me.weight_mult.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("wp_mult = " & Me.wp_mult.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("af_factor = " & Me.af_factor.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ar_factor = " & Me.ar_factor.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("round_area_ratio = " & Me.round_area_ratio.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flat_area_ratio = " & Me.flat_area_ratio.ToString.FormatDBValue)

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
        Dim otherToCompare As PoleSection = TryCast(other, PoleSection)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.section_id.CheckChange(otherToCompare.section_id, changes, categoryName, "Section Id"), Equals, False)
        'Equals = If(Me.pole_id.CheckChange(otherToCompare.pole_id, changes, categoryName, "Pole Id"), Equals, False)
        Equals = If(Me.local_section_id.CheckChange(otherToCompare.local_section_id, changes, categoryName, "Local Section Id"), Equals, False)
        Equals = If(Me.elev_bot.CheckChange(otherToCompare.elev_bot, changes, categoryName, "Elev Bot"), Equals, False)
        Equals = If(Me.elev_top.CheckChange(otherToCompare.elev_top, changes, categoryName, "Elev Top"), Equals, False)
        Equals = If(Me.length_section.CheckChange(otherToCompare.length_section, changes, categoryName, "Length Section"), Equals, False)
        Equals = If(Me.length_splice.CheckChange(otherToCompare.length_splice, changes, categoryName, "Length Splice"), Equals, False)
        Equals = If(Me.num_sides.CheckChange(otherToCompare.num_sides, changes, categoryName, "Num Sides"), Equals, False)
        Equals = If(Me.diam_bot.CheckChange(otherToCompare.diam_bot, changes, categoryName, "Diam Bot"), Equals, False)
        Equals = If(Me.diam_top.CheckChange(otherToCompare.diam_top, changes, categoryName, "Diam Top"), Equals, False)
        Equals = If(Me.wall_thickness.CheckChange(otherToCompare.wall_thickness, changes, categoryName, "Wall Thickness"), Equals, False)
        Equals = If(Me.bend_radius.CheckChange(otherToCompare.bend_radius, changes, categoryName, "Bend Radius"), Equals, False)
        'Equals = If(Me.matl_id.CheckChange(otherToCompare.matl_id, changes, categoryName, "Matl Id"), Equals, False)
        Equals = If(Me.local_matl_id.CheckChange(otherToCompare.local_matl_id, changes, categoryName, "Local Matl Id"), Equals, False)
        Equals = If(Me.pole_type.CheckChange(otherToCompare.pole_type, changes, categoryName, "Pole Type"), Equals, False)
        Equals = If(Me.section_name.CheckChange(otherToCompare.section_name, changes, categoryName, "Section Name"), Equals, False)
        Equals = If(Me.socket_length.CheckChange(otherToCompare.socket_length, changes, categoryName, "Socket Length"), Equals, False)
        Equals = If(Me.weight_mult.CheckChange(otherToCompare.weight_mult, changes, categoryName, "Weight Mult"), Equals, False)
        Equals = If(Me.wp_mult.CheckChange(otherToCompare.wp_mult, changes, categoryName, "Wp Mult"), Equals, False)
        Equals = If(Me.af_factor.CheckChange(otherToCompare.af_factor, changes, categoryName, "Af Factor"), Equals, False)
        Equals = If(Me.ar_factor.CheckChange(otherToCompare.ar_factor, changes, categoryName, "Ar Factor"), Equals, False)
        Equals = If(Me.round_area_ratio.CheckChange(otherToCompare.round_area_ratio, changes, categoryName, "Round Area Ratio"), Equals, False)
        Equals = If(Me.flat_area_ratio.CheckChange(otherToCompare.flat_area_ratio, changes, categoryName, "Flat Area Ratio"), Equals, False)

    End Function
#End Region

End Class

Partial Public Class PoleReinfSection
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Reinforced Sections"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.reinforced_sections"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\3 Reinf Section (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Section_INSERT

        If IsSomething(Me.matl_id) And Me.matl_id <> 0 Then
            SQLInsert = SQLInsert.Replace("[MATL ID]", Me.matl_id.ToString.FormatDBValue)
        Else
            SQLInsert = SQLInsert.Replace("[MATL ID]", "NULL")

            Dim ParentPole As Pole = TryCast(Me.Parent, Pole)
            If IsSomething(ParentPole) Then
                For Each dbrow As PoleMatlProp In ParentPole.matls
                    If Me.local_matl_id = dbrow.local_matl_id Then 'And Me.local_matl_id > 17 Then 'Matching, Non-Standard Materials
                        SQLInsert = SQLInsert.Replace("--[MATL DB SUBQUERY]", dbrow.SQLInsert)
                    End If
                Next
            End If

        End If

        SQLInsert = SQLInsert.Replace("[REINF SECTION VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[REINF SECTION FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\3 Reinf Section (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Section_UPDATE

        If IsSomething(Me.matl_id) And Me.matl_id <> 0 Then
            SQLUpdate = SQLUpdate.Replace("[MATL ID]", Me.matl_id.ToString.FormatDBValue)
        Else
            SQLUpdate = SQLUpdate.Replace("[MATL ID]", "NULL")

            Dim ParentPole As Pole = TryCast(Me.Parent, Pole)
            If IsSomething(ParentPole) Then
                For Each dbrow As PoleMatlProp In ParentPole.matls
                    If Me.local_matl_id = dbrow.local_matl_id Then
                        SQLUpdate = SQLUpdate.Replace("--[MATL DB SUBQUERY]", dbrow.SQLInsert)
                    End If
                Next
            End If

        End If

        SQLUpdate = SQLUpdate.Replace("[ID]", Me.section_id.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\3 Reinf Section (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Section_DELETE

        SQLDelete = SQLDelete.Replace("[ID]", Me.section_id.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete
    End Function
#End Region

#Region "Define"
    Private _section_id As Integer?
    Private _pole_id As Integer?
    Private _local_section_id As Integer?
    Private _elev_bot As Double?
    Private _elev_top As Double?
    Private _length_section As Double?
    Private _length_splice As Double?
    Private _num_sides As Integer?
    Private _diam_bot As Double?
    Private _diam_top As Double?
    Private _wall_thickness As Double?
    Private _bend_radius As Double?
    Private _matl_id As Integer?
    Private _local_matl_id As Integer?
    Private _pole_type As String
    Private _weight_mult As Double?
    Private _section_name As String
    Private _socket_length As Double?
    Private _wp_mult As Double?
    Private _af_factor As Double?
    Private _ar_factor As Double?
    Private _round_area_ratio As Double?
    Private _flat_area_ratio As Double?
    'Public Property matls As New List(Of PoleMatlProp)

    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Section Id")>
    Public Property section_id() As Integer?
        Get
            Return Me._section_id
        End Get
        Set
            Me._section_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Pole Id")>
    Public Property pole_id() As Integer?
        Get
            Return Me._pole_id
        End Get
        Set
            Me._pole_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Local Section Id")>
    Public Property local_section_id() As Integer?
        Get
            Return Me._local_section_id
        End Get
        Set
            Me._local_section_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Elev Bot")>
    Public Property elev_bot() As Double?
        Get
            Return Me._elev_bot
        End Get
        Set
            Me._elev_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Elev Top")>
    Public Property elev_top() As Double?
        Get
            Return Me._elev_top
        End Get
        Set
            Me._elev_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Length Section")>
    Public Property length_section() As Double?
        Get
            Return Me._length_section
        End Get
        Set
            Me._length_section = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Length Splice")>
    Public Property length_splice() As Double?
        Get
            Return Me._length_splice
        End Get
        Set
            Me._length_splice = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Num Sides")>
    Public Property num_sides() As Integer?
        Get
            Return Me._num_sides
        End Get
        Set
            Me._num_sides = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Diam Bot")>
    Public Property diam_bot() As Double?
        Get
            Return Me._diam_bot
        End Get
        Set
            Me._diam_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Diam Top")>
    Public Property diam_top() As Double?
        Get
            Return Me._diam_top
        End Get
        Set
            Me._diam_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Wall Thickness")>
    Public Property wall_thickness() As Double?
        Get
            Return Me._wall_thickness
        End Get
        Set
            Me._wall_thickness = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Bend Radius")>
    Public Property bend_radius() As Double?
        Get
            Return Me._bend_radius
        End Get
        Set
            Me._bend_radius = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Matl Id")>
    Public Property matl_id() As Integer?
        Get
            Return Me._matl_id
        End Get
        Set
            Me._matl_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Local Matl Id")>
    Public Property local_matl_id() As Integer?
        Get
            Return Me._local_matl_id
        End Get
        Set
            Me._local_matl_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Pole Type")>
    Public Property pole_type() As String
        Get
            Return Me._pole_type
        End Get
        Set
            Me._pole_type = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Weight Mult")>
    Public Property weight_mult() As Double?
        Get
            Return Me._weight_mult
        End Get
        Set
            Me._weight_mult = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Section Name")>
    Public Property section_name() As String
        Get
            Return Me._section_name
        End Get
        Set
            Me._section_name = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Socket Length")>
    Public Property socket_length() As Double?
        Get
            Return Me._socket_length
        End Get
        Set
            Me._socket_length = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Wp Mult")>
    Public Property wp_mult() As Double?
        Get
            Return Me._wp_mult
        End Get
        Set
            Me._wp_mult = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Af Factor")>
    Public Property af_factor() As Double?
        Get
            Return Me._af_factor
        End Get
        Set
            Me._af_factor = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Ar Factor")>
    Public Property ar_factor() As Double?
        Get
            Return Me._ar_factor
        End Get
        Set
            Me._ar_factor = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Round Area Ratio")>
    Public Property round_area_ratio() As Double?
        Get
            Return Me._round_area_ratio
        End Get
        Set
            Me._round_area_ratio = Value
        End Set
    End Property
    <Category("CCIpole Reinf Sections"), Description(""), DisplayName("Flat Area Ratio")>
    Public Property flat_area_ratio() As Double?
        Get
            Return Me._flat_area_ratio
        End Get
        Set
            Me._flat_area_ratio = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByVal Parent As EDSObject = Nothing) ', ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row

        Me.section_id = DBtoNullableInt(dr.Item("ID"))
        Me.pole_id = DBtoNullableInt(dr.Item("pole_id"))
        Me.local_section_id = DBtoNullableInt(dr.Item("local_section_id"))
        Me.elev_bot = DBtoNullableDbl(dr.Item("elev_bot"))
        Me.elev_top = DBtoNullableDbl(dr.Item("elev_top"))
        Me.length_section = DBtoNullableDbl(dr.Item("length_section"))
        Me.length_splice = DBtoNullableDbl(dr.Item("length_splice"))
        Me.num_sides = DBtoNullableInt(dr.Item("num_sides"))
        Me.diam_bot = DBtoNullableDbl(dr.Item("diam_bot"))
        Me.diam_top = DBtoNullableDbl(dr.Item("diam_top"))
        Me.wall_thickness = DBtoNullableDbl(dr.Item("wall_thickness"))
        Me.bend_radius = If(DBtoStr(dr.Item("bend_radius")) = "Auto", -1, DBtoNullableDbl(dr.Item("bend_radius")))
        Me.matl_id = DBtoNullableInt(dr.Item("matl_id"))
        Me.local_matl_id = DBtoNullableInt(dr.Item("local_matl_id"))
        Me.pole_type = DBtoStr(dr.Item("pole_type"))
        Me.weight_mult = DBtoNullableDbl(dr.Item("weight_mult"))
        Me.section_name = DBtoStr(dr.Item("section_name"))
        Me.socket_length = DBtoNullableDbl(dr.Item("socket_length"))
        Me.wp_mult = DBtoNullableDbl(dr.Item("wp_mult"))
        Me.af_factor = DBtoNullableDbl(dr.Item("af_factor"))
        Me.ar_factor = DBtoNullableDbl(dr.Item("ar_factor"))
        Me.round_area_ratio = DBtoNullableDbl(dr.Item("round_area_ratio"))
        Me.flat_area_ratio = DBtoNullableDbl(dr.Item("flat_area_ratio"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.section_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") '(Me.pole_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_section_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.length_section.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.length_splice.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.num_sides.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.diam_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.diam_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.wall_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bend_radius.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel4ID") '(Me.matl_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_matl_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weight_mult.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.section_name.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.socket_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.wp_mult.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.af_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ar_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.round_area_ratio.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.flat_area_ratio.ToString.FormatDBValue)


        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("length_section")
        SQLInsertFields = SQLInsertFields.AddtoDBString("length_splice")
        SQLInsertFields = SQLInsertFields.AddtoDBString("num_sides")
        SQLInsertFields = SQLInsertFields.AddtoDBString("diam_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("diam_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("wall_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bend_radius")
        SQLInsertFields = SQLInsertFields.AddtoDBString("matl_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_matl_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weight_mult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("section_name")
        SQLInsertFields = SQLInsertFields.AddtoDBString("socket_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("wp_mult")
        SQLInsertFields = SQLInsertFields.AddtoDBString("af_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ar_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("round_area_ratio")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flat_area_ratio")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        Dim ParentPole As Pole = TryCast(Me.Parent, Pole)

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("section_id = " & Me.section_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_id = " & ParentPole.pole_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_section_id = " & Me.local_section_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_bot = " & Me.elev_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_top = " & Me.elev_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("length_section = " & Me.length_section.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("length_splice = " & Me.length_splice.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("num_sides = " & Me.num_sides.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("diam_bot = " & Me.diam_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("diam_top = " & Me.diam_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("wall_thickness = " & Me.wall_thickness.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bend_radius = " & Me.bend_radius.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("matl_id = @SubLevel4ID") '& Me.matl_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_matl_id = " & Me.local_matl_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_type = " & Me.pole_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weight_mult = " & Me.weight_mult.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("section_name = " & Me.section_name.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("socket_length = " & Me.socket_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("wp_mult = " & Me.wp_mult.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("af_factor = " & Me.af_factor.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ar_factor = " & Me.ar_factor.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("round_area_ratio = " & Me.round_area_ratio.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flat_area_ratio = " & Me.flat_area_ratio.ToString.FormatDBValue)

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
        Dim otherToCompare As PoleReinfSection = TryCast(other, PoleReinfSection)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.section_id.CheckChange(otherToCompare.section_id, changes, categoryName, "Section Id"), Equals, False)
        'Equals = If(Me.pole_id.CheckChange(otherToCompare.pole_id, changes, categoryName, "Pole Id"), Equals, False)
        Equals = If(Me.local_section_id.CheckChange(otherToCompare.local_section_id, changes, categoryName, "Local Section Id"), Equals, False)
        Equals = If(Me.elev_bot.CheckChange(otherToCompare.elev_bot, changes, categoryName, "Elev Bot"), Equals, False)
        Equals = If(Me.elev_top.CheckChange(otherToCompare.elev_top, changes, categoryName, "Elev Top"), Equals, False)
        Equals = If(Me.length_section.CheckChange(otherToCompare.length_section, changes, categoryName, "Length Section"), Equals, False)
        Equals = If(Me.length_splice.CheckChange(otherToCompare.length_splice, changes, categoryName, "Length Splice"), Equals, False)
        Equals = If(Me.num_sides.CheckChange(otherToCompare.num_sides, changes, categoryName, "Num Sides"), Equals, False)
        Equals = If(Me.diam_bot.CheckChange(otherToCompare.diam_bot, changes, categoryName, "Diam Bot"), Equals, False)
        Equals = If(Me.diam_top.CheckChange(otherToCompare.diam_top, changes, categoryName, "Diam Top"), Equals, False)
        Equals = If(Me.wall_thickness.CheckChange(otherToCompare.wall_thickness, changes, categoryName, "Wall Thickness"), Equals, False)
        Equals = If(Me.bend_radius.CheckChange(otherToCompare.bend_radius, changes, categoryName, "Bend Radius"), Equals, False)
        'Equals = If(Me.matl_id.CheckChange(otherToCompare.matl_id, changes, categoryName, "Matl Id"), Equals, False)
        Equals = If(Me.local_matl_id.CheckChange(otherToCompare.local_matl_id, changes, categoryName, "Local Matl Id"), Equals, False)
        Equals = If(Me.pole_type.CheckChange(otherToCompare.pole_type, changes, categoryName, "Pole Type"), Equals, False)
        Equals = If(Me.weight_mult.CheckChange(otherToCompare.weight_mult, changes, categoryName, "Weight Mult"), Equals, False)
        Equals = If(Me.section_name.CheckChange(otherToCompare.section_name, changes, categoryName, "Section Name"), Equals, False)
        Equals = If(Me.socket_length.CheckChange(otherToCompare.socket_length, changes, categoryName, "Socket Length"), Equals, False)
        Equals = If(Me.wp_mult.CheckChange(otherToCompare.wp_mult, changes, categoryName, "Wp Mult"), Equals, False)
        Equals = If(Me.af_factor.CheckChange(otherToCompare.af_factor, changes, categoryName, "Af Factor"), Equals, False)
        Equals = If(Me.ar_factor.CheckChange(otherToCompare.ar_factor, changes, categoryName, "Ar Factor"), Equals, False)
        Equals = If(Me.round_area_ratio.CheckChange(otherToCompare.round_area_ratio, changes, categoryName, "Round Area Ratio"), Equals, False)
        Equals = If(Me.flat_area_ratio.CheckChange(otherToCompare.flat_area_ratio, changes, categoryName, "Flat Area Ratio"), Equals, False)

    End Function
#End Region

End Class

Partial Public Class PoleReinfGroup
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Reinforcement Groups"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.reinforcements"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\4 Reinf Group (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Group_INSERT

        If IsSomething(Me.reinf_id) And Me.reinf_id <> 0 Then 'Default Reinf Used - no need to pull in subqueries in order to create Reinf Group
            SQLInsert = SQLInsert.Replace("--[REINF DB SUBQUERY]", "SET @SubLevel2ID = " & Me.reinf_id.ToString.FormatDBValue)
        Else ' Custom Reinf/Matl/Bolt used
            Dim ParentPole As Pole = TryCast(Me.Parent, Pole)
            If IsSomething(ParentPole) Then
                For Each dbrow As PoleReinfProp In ParentPole.reinfs
                    If Me.local_reinf_id = dbrow.local_reinf_id Then
                        SQLInsert = SQLInsert.Replace("--[REINF DB SUBQUERY]", dbrow.SQLInsert)
                    End If
                Next
            End If

        End If

        SQLInsert = SQLInsert.Replace("[REINF GROUP VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[REINF GROUP FIELDS]", Me.SQLInsertFields)

        For Each detailrow As PoleReinfDetail In reinf_ids
            If Me.local_group_id = detailrow.local_group_id Then
                SQLInsert = SQLInsert.Replace("--[REINF DETAIL SUBQUERY]", detailrow.SQLInsert)
            End If
        Next

        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\4 Reinf Group (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Group_UPDATE

        If IsSomething(Me.reinf_id) And Me.reinf_id <> 0 Then 'Default Reinf Used - no need to pull in subqueries in order to create Reinf Group
            SQLUpdate = SQLUpdate.Replace("--[REINF DB SUBQUERY]", "SET @SubLevel2ID = " & Me.reinf_id.ToString.FormatDBValue)
        Else ' Custom Reinf/Matl/Bolt used
            Dim ParentPole As Pole = TryCast(Me.Parent, Pole)
            If IsSomething(ParentPole) Then
                For Each dbrow As PoleReinfProp In ParentPole.reinfs
                    If Me.local_reinf_id = dbrow.local_reinf_id Then
                        SQLUpdate = SQLUpdate.Replace("--[REINF DB SUBQUERY]", dbrow.SQLInsert)
                    End If
                Next
            End If

        End If

        SQLUpdate = SQLUpdate.Replace("[ID]", Me.group_id.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)

        For Each detailrow As PoleReinfDetail In reinf_ids
            If IsSomething(detailrow.reinforcement_id) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                If IsSomething(detailrow.local_reinforcement_id) Then
                    SQLUpdate = SQLUpdate.Replace("--[REINF DETAIL SUBQUERY]", detailrow.SQLUpdate)
                Else
                    SQLUpdate = SQLUpdate.Replace("--[REINF DETAIL SUBQUERY]", detailrow.SQLDelete)
                End If
            Else
                SQLUpdate = SQLUpdate.Replace("--[REINF DETAIL SUBQUERY]", detailrow.SQLInsert)
            End If
        Next

        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\4 Reinf Group (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Group_DELETE

        SQLDelete = SQLDelete.Replace("[ID]", Me.group_id.ToString.FormatDBValue)

        SQLDelete = SQLDelete.Replace("--[REINF DETAIL SUBQUERY]", "DELETE From pole.reinforcement_details Where group_id = " & Me.group_id.ToString.FormatDBValue)
        'For Each detailrow As PoleReinfDetail In reinf_ids
        '    If Me.local_group_id = detailrow.local_group_id Then
        '        SQLDelete = SQLDelete.Replace("--[REINF DETAIL SUBQUERY]", detailrow.SQLDelete)
        '    End If
        'Next

        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete
    End Function
#End Region

#Region "Define"
    Private _group_id As Integer?
    Private _pole_id As Integer?
    Private _local_group_id As Integer?
    Private _elev_bot_actual As Double?
    Private _elev_bot_eff As Double?
    Private _elev_top_actual As Double?
    Private _elev_top_eff As Double?
    Private _reinf_id As Integer?
    Private _local_reinf_id As Integer?
    Private _qty As Integer?
    'Public Property reinfs As New List(Of PoleReinfProp)
    Public Property reinf_ids As New List(Of PoleReinfDetail)

    <Category("CCIpole Reinfs"), Description(""), DisplayName("Group Id")>
    Public Property group_id() As Integer?
        Get
            Return Me._group_id
        End Get
        Set
            Me._group_id = Value
        End Set
    End Property
    <Category("CCIpole Reinfs"), Description(""), DisplayName("Pole Id")>
    Public Property pole_id() As Integer?
        Get
            Return Me._pole_id
        End Get
        Set
            Me._pole_id = Value
        End Set
    End Property
    <Category("CCIpole Reinfs"), Description(""), DisplayName("Local Group Id")>
    Public Property local_group_id() As Integer?
        Get
            Return Me._local_group_id
        End Get
        Set
            Me._local_group_id = Value
        End Set
    End Property
    <Category("CCIpole Reinfs"), Description(""), DisplayName("Elev Bot Actual")>
    Public Property elev_bot_actual() As Double?
        Get
            Return Me._elev_bot_actual
        End Get
        Set
            Me._elev_bot_actual = Value
        End Set
    End Property
    <Category("CCIpole Reinfs"), Description(""), DisplayName("Elev Bot Eff")>
    Public Property elev_bot_eff() As Double?
        Get
            Return Me._elev_bot_eff
        End Get
        Set
            Me._elev_bot_eff = Value
        End Set
    End Property
    <Category("CCIpole Reinfs"), Description(""), DisplayName("Elev Top Actual")>
    Public Property elev_top_actual() As Double?
        Get
            Return Me._elev_top_actual
        End Get
        Set
            Me._elev_top_actual = Value
        End Set
    End Property
    <Category("CCIpole Reinfs"), Description(""), DisplayName("Elev Top Eff")>
    Public Property elev_top_eff() As Double?
        Get
            Return Me._elev_top_eff
        End Get
        Set
            Me._elev_top_eff = Value
        End Set
    End Property
    <Category("CCIpole Reinfs"), Description(""), DisplayName("Reinf Id")>
    Public Property reinf_id() As Integer?
        Get
            Return Me._reinf_id
        End Get
        Set
            Me._reinf_id = Value
        End Set
    End Property
    <Category("CCIpole Reinfs"), Description(""), DisplayName("Local Reinf Id")>
    Public Property local_reinf_id() As Integer?
        Get
            Return Me._local_reinf_id
        End Get
        Set
            Me._local_reinf_id = Value
        End Set
    End Property
    <Category("CCIpole Reinfs"), Description(""), DisplayName("Qty")>
    Public Property qty() As Integer?
        Get
            Return Me._qty
        End Get
        Set
            Me._qty = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByVal Parent As EDSObject = Nothing) ', ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row

        Me.group_id = DBtoNullableInt(dr.Item("ID"))
        Me.pole_id = DBtoNullableInt(dr.Item("pole_id"))
        Me.local_group_id = DBtoNullableInt(dr.Item("local_group_id"))
        Me.elev_bot_actual = DBtoNullableDbl(dr.Item("elev_bot_actual"))
        Me.elev_bot_eff = DBtoNullableDbl(dr.Item("elev_bot_eff"))
        Me.elev_top_actual = DBtoNullableDbl(dr.Item("elev_top_actual"))
        Me.elev_top_eff = DBtoNullableDbl(dr.Item("elev_top_eff"))
        Me.reinf_id = DBtoNullableInt(dr.Item("reinf_id"))
        Me.local_reinf_id = DBtoNullableInt(dr.Item("local_reinf_id"))
        Me.qty = DBtoNullableInt(dr.Item("qty"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID") '(Me.group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") '(Me.pole_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_bot_actual.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_bot_eff.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_top_actual.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_top_eff.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID") '(Me.reinf_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_reinf_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.qty.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_bot_actual")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_bot_eff")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_top_actual")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_top_eff")
        SQLInsertFields = SQLInsertFields.AddtoDBString("reinf_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_reinf_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("qty")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        Dim ParentPole As Pole = TryCast(Me.Parent, Pole)

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("group_id = " & Me.group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_id = " & ParentPole.pole_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_group_id = " & Me.local_group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_bot_actual = " & Me.elev_bot_actual.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_bot_eff = " & Me.elev_bot_eff.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_top_actual = " & Me.elev_top_actual.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_top_eff = " & Me.elev_top_eff.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("reinf_id = @SubLevel2ID") '& Me.reinf_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_reinf_id = " & Me.local_reinf_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("qty = " & Me.qty.ToString.FormatDBValue)

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
        Dim otherToCompare As PoleReinfGroup = TryCast(other, PoleReinfGroup)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.group_id.CheckChange(otherToCompare.group_id, changes, categoryName, "Group Id"), Equals, False)
        'Equals = If(Me.pole_id.CheckChange(otherToCompare.pole_id, changes, categoryName, "Pole Id"), Equals, False)
        Equals = If(Me.local_group_id.CheckChange(otherToCompare.local_group_id, changes, categoryName, "Local Group Id"), Equals, False)
        Equals = If(Me.elev_bot_actual.CheckChange(otherToCompare.elev_bot_actual, changes, categoryName, "Elev Bot Actual"), Equals, False)
        Equals = If(Me.elev_bot_eff.CheckChange(otherToCompare.elev_bot_eff, changes, categoryName, "Elev Bot Eff"), Equals, False)
        Equals = If(Me.elev_top_actual.CheckChange(otherToCompare.elev_top_actual, changes, categoryName, "Elev Top Actual"), Equals, False)
        Equals = If(Me.elev_top_eff.CheckChange(otherToCompare.elev_top_eff, changes, categoryName, "Elev Top Eff"), Equals, False)
        'Equals = If(Me.reinf_id.CheckChange(otherToCompare.reinf_id, changes, categoryName, "Reinf Id"), Equals, False)
        Equals = If(Me.local_reinf_id.CheckChange(otherToCompare.local_reinf_id, changes, categoryName, "Local Reinf Id"), Equals, False)
        Equals = If(Me.qty.CheckChange(otherToCompare.qty, changes, categoryName, "Qty"), Equals, False)

        Equals = If(Me.reinf_ids.CheckChange(otherToCompare.reinf_ids, changes, categoryName, "Reinforcement Details"), Equals, False)

    End Function
#End Region

End Class

Partial Public Class PoleReinfDetail
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Reinforcement Details"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.reinforcement_details"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\5 Reinf Detail (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Detail_INSERT

        SQLInsert = SQLInsert.Replace("[REINF DETAIL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[REINF DETAIL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\5 Reinf Detail (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Detail_UPDATE

        SQLUpdate = SQLUpdate.Replace("[ID]", Me.reinforcement_id.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\5 Reinf Detail (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Detail_DELETE

        SQLDelete = SQLDelete.Replace("[ID]", Me.reinforcement_id.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete
    End Function
#End Region

#Region "Define"
    Private _reinforcement_id As Integer?
    Private _group_id As Integer?
    Private _local_group_id As Integer?
    Private _local_reinforcement_id As Integer?
    Private _pole_flat As Integer?
    Private _horizontal_offset As Double?
    Private _rotation As Double?
    Private _note As String

    <Category("CCIpole Reinf Details"), Description(""), DisplayName("Reinforcment Id")>
    Public Property reinforcement_id() As Integer?
        Get
            Return Me._reinforcement_id
        End Get
        Set
            Me._reinforcement_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf Details"), Description(""), DisplayName("Group Id")>
    Public Property group_id() As Integer?
        Get
            Return Me._group_id
        End Get
        Set
            Me._group_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf Details"), Description(""), DisplayName("Local Group Id")>
    Public Property local_group_id() As Integer?
        Get
            Return Me._local_group_id
        End Get
        Set
            Me._local_group_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf Details"), Description(""), DisplayName("Local Reinf Id")>
    Public Property local_reinforcement_id() As Integer?
        Get
            Return Me._local_reinforcement_id
        End Get
        Set
            Me._local_reinforcement_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf Details"), Description(""), DisplayName("Pole Flat")>
    Public Property pole_flat() As Integer?
        Get
            Return Me._pole_flat
        End Get
        Set
            Me._pole_flat = Value
        End Set
    End Property
    <Category("CCIpole Reinf Details"), Description(""), DisplayName("Horizontal Offset")>
    Public Property horizontal_offset() As Double?
        Get
            Return Me._horizontal_offset
        End Get
        Set
            Me._horizontal_offset = Value
        End Set
    End Property
    <Category("CCIpole Reinf Details"), Description(""), DisplayName("Rotation")>
    Public Property rotation() As Double?
        Get
            Return Me._rotation
        End Get
        Set
            Me._rotation = Value
        End Set
    End Property
    <Category("CCIpole Reinf Details"), Description(""), DisplayName("Note")>
    Public Property note() As String
        Get
            Return Me._note
        End Get
        Set
            Me._note = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByRef Parent As PoleReinfGroup = Nothing) ', ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(DirectCast(Parent, EDSObject))
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row

        Me.reinforcement_id = DBtoNullableInt(dr.Item("ID"))
        Me.group_id = DBtoNullableInt(dr.Item("group_id"))
        Me.local_group_id = DBtoNullableInt(dr.Item("local_group_id"))
        Me.local_reinforcement_id = DBtoNullableInt(dr.Item("local_reinforcement_id"))
        Me.pole_flat = DBtoNullableInt(dr.Item("pole_flat"))
        Me.horizontal_offset = DBtoNullableDbl(dr.Item("horizontal_offset"))
        Me.rotation = DBtoNullableDbl(dr.Item("rotation"))
        Me.note = DBtoStr(dr.Item("note"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reinforcement_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID") '(Me.group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_reinforcement_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_flat.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.horizontal_offset.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rotation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.note.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("reinforcement_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_reinforcement_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_flat")
        SQLInsertFields = SQLInsertFields.AddtoDBString("horizontal_offset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rotation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("note")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        Dim ParentGroup As PoleReinfGroup = TryCast(Me.Parent, PoleReinfGroup)

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("reinforcement_id = " & Me.reinforcement_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("group_id = " & ParentGroup.group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_group_id = " & Me.local_group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_reinforcement_id = " & Me.local_reinforcement_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_flat = " & Me.pole_flat.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("horizontal_offset = " & Me.horizontal_offset.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rotation = " & Me.rotation.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("note = " & Me.note.ToString.FormatDBValue)

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
        Dim otherToCompare As PoleReinfDetail = TryCast(other, PoleReinfDetail)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.reinforcement_id.CheckChange(otherToCompare.reinforcement_id, changes, categoryName, "Reinforcment Id"), Equals, False)
        'Equals = If(Me.group_id.CheckChange(otherToCompare.group_id, changes, categoryName, "Group Id"), Equals, False)
        Equals = If(Me.local_group_id.CheckChange(otherToCompare.local_group_id, changes, categoryName, "Local Group Id"), Equals, False)
        Equals = If(Me.local_reinforcement_id.CheckChange(otherToCompare.local_reinforcement_id, changes, categoryName, "Local Reinforcement Id"), Equals, False)
        Equals = If(Me.pole_flat.CheckChange(otherToCompare.pole_flat, changes, categoryName, "Pole Flat"), Equals, False)
        Equals = If(Me.horizontal_offset.CheckChange(otherToCompare.horizontal_offset, changes, categoryName, "Horizontal Offset"), Equals, False)
        Equals = If(Me.rotation.CheckChange(otherToCompare.rotation, changes, categoryName, "Rotation"), Equals, False)
        Equals = If(Me.note.CheckChange(otherToCompare.note, changes, categoryName, "Note"), Equals, False)

    End Function
#End Region

End Class

Partial Public Class PoleIntGroup
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Interference Groups"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.interferences"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\6 Int Group (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_Int_Group_INSERT

        SQLInsert = SQLInsert.Replace("[INT GROUP VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[INT GROUP FIELDS]", Me.SQLInsertFields)

        For Each detailrow As PoleIntDetail In int_ids
            If Me.local_group_id = detailrow.local_group_id Then
                SQLInsert = SQLInsert.Replace("--[INT DETAIL SUBQUERY]", detailrow.SQLInsert)
            End If
        Next

        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\6 Int Group (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIpole_Int_Group_UPDATE

        SQLUpdate = SQLUpdate.Replace("[ID]", Me.group_id.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)

        For Each detailrow As PoleIntDetail In int_ids
            If IsSomething(detailrow.interference_id) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                If IsSomething(detailrow.local_interference_id) Then
                    SQLUpdate = SQLUpdate.Replace("--[INT DETAIL SUBQUERY]", detailrow.SQLUpdate)
                Else
                    SQLUpdate = SQLUpdate.Replace("--[INT DETAIL SUBQUERY]", detailrow.SQLDelete)
                End If
            Else
                SQLUpdate = SQLUpdate.Replace("--[INT DETAIL SUBQUERY]", detailrow.SQLInsert)
            End If
        Next

        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\6 Int Group (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIpole_Int_Group_DELETE

        SQLDelete = SQLDelete.Replace("[ID]", Me.group_id.ToString.FormatDBValue)

        SQLDelete = SQLDelete.Replace("--[INT DETAIL SUBQUERY]", "DELETE From pole.interference_details Where group_id = " & Me.group_id.ToString.FormatDBValue)
        'For Each detailrow As PoleIntDetail In int_ids
        '    If Me.local_group_id = detailrow.local_group_id Then
        '        SQLDelete = SQLDelete.Replace("--[INT DETAIL SUBQUERY]", detailrow.SQLDelete)
        '    End If
        'Next

        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete
    End Function
#End Region

#Region "Define"
    Private _group_id As Integer?
    Private _pole_id As Integer?
    Private _local_group_id As Integer?
    Private _elev_bot As Double?
    Private _elev_top As Double?
    Private _width As Double?
    Private _description As String
    Private _qty As Integer?
    Public Property int_ids As New List(Of PoleIntDetail)

    <Category("CCIpole Ints"), Description(""), DisplayName("Group Id")>
    Public Property group_id() As Integer?
        Get
            Return Me._group_id
        End Get
        Set
            Me._group_id = Value
        End Set
    End Property
    <Category("CCIpole Ints"), Description(""), DisplayName("Pole Id")>
    Public Property pole_id() As Integer?
        Get
            Return Me._pole_id
        End Get
        Set
            Me._pole_id = Value
        End Set
    End Property
    <Category("CCIpole Ints"), Description(""), DisplayName("Local Group Id")>
    Public Property local_group_id() As Integer?
        Get
            Return Me._local_group_id
        End Get
        Set
            Me._local_group_id = Value
        End Set
    End Property
    <Category("CCIpole Ints"), Description(""), DisplayName("Elev Bot")>
    Public Property elev_bot() As Double?
        Get
            Return Me._elev_bot
        End Get
        Set
            Me._elev_bot = Value
        End Set
    End Property
    <Category("CCIpole Ints"), Description(""), DisplayName("Elev Top")>
    Public Property elev_top() As Double?
        Get
            Return Me._elev_top
        End Get
        Set
            Me._elev_top = Value
        End Set
    End Property
    <Category("CCIpole Ints"), Description(""), DisplayName("Width")>
    Public Property width() As Double?
        Get
            Return Me._width
        End Get
        Set
            Me._width = Value
        End Set
    End Property
    <Category("CCIpole Ints"), Description(""), DisplayName("Description")>
    Public Property description() As String
        Get
            Return Me._description
        End Get
        Set
            Me._description = Value
        End Set
    End Property
    <Category("CCIpole Ints"), Description(""), DisplayName("Qty")>
    Public Property qty() As Integer?
        Get
            Return Me._qty
        End Get
        Set
            Me._qty = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByVal Parent As EDSObject = Nothing) ', ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row

        Me.group_id = DBtoNullableInt(dr.Item("ID"))
        Me.pole_id = DBtoNullableInt(dr.Item("pole_id"))
        Me.local_group_id = DBtoNullableInt(dr.Item("local_group_id"))
        Me.elev_bot = DBtoNullableDbl(dr.Item("elev_bot"))
        Me.elev_top = DBtoNullableDbl(dr.Item("elev_top"))
        Me.width = DBtoNullableDbl(dr.Item("width"))
        Me.description = DBtoStr(dr.Item("description"))
        Me.qty = DBtoNullableInt(dr.Item("qty"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") '(Me.pole_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elev_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.description.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.qty.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elev_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("description")
        SQLInsertFields = SQLInsertFields.AddtoDBString("qty")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        Dim ParentPole As Pole = TryCast(Me.Parent, Pole)

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("group_id = " & Me.group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_id = " & ParentPole.pole_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_group_id = " & Me.local_group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_bot = " & Me.elev_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elev_top = " & Me.elev_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("width = " & Me.width.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("description = " & Me.description.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("qty = " & Me.qty.ToString.FormatDBValue)

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
        Dim otherToCompare As PoleIntGroup = TryCast(other, PoleIntGroup)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.group_id.CheckChange(otherToCompare.group_id, changes, categoryName, "Group Id"), Equals, False)
        'Equals = If(Me.pole_id.CheckChange(otherToCompare.pole_id, changes, categoryName, "Pole Id"), Equals, False)
        Equals = If(Me.local_group_id.CheckChange(otherToCompare.local_group_id, changes, categoryName, "Local Group Id"), Equals, False)
        Equals = If(Me.elev_bot.CheckChange(otherToCompare.elev_bot, changes, categoryName, "Elev Bot"), Equals, False)
        Equals = If(Me.elev_top.CheckChange(otherToCompare.elev_top, changes, categoryName, "Elev Top"), Equals, False)
        Equals = If(Me.width.CheckChange(otherToCompare.width, changes, categoryName, "Width"), Equals, False)
        Equals = If(Me.description.CheckChange(otherToCompare.description, changes, categoryName, "Description"), Equals, False)
        Equals = If(Me.qty.CheckChange(otherToCompare.qty, changes, categoryName, "Qty"), Equals, False)

        Equals = If(Me.int_ids.CheckChange(otherToCompare.int_ids, changes, categoryName, "Interference Details"), Equals, False)

    End Function
#End Region

End Class

Partial Public Class PoleIntDetail
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Interference Details"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.interference_details"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\7 Int Detail (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_Int_Detail_INSERT

        SQLInsert = SQLInsert.Replace("[INT DETAIL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[INT DETAIL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\7 Int Detail (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIpole_Int_Detail_UPDATE

        SQLUpdate = SQLUpdate.Replace("[ID]", Me.interference_id.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\7 Int Detail (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIpole_Int_Detail_DELETE

        SQLDelete = SQLDelete.Replace("[ID]", Me.interference_id.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete
    End Function
#End Region

#Region "Define"
    Private _interference_id As Integer?
    Private _group_id As Integer?
    Private _local_group_id As Integer?
    Private _local_interference_id As Integer?
    Private _pole_flat As Integer?
    Private _horizontal_offset As Double?
    Private _rotation As Double?
    Private _note As String

    <Category("CCIpole Int Details"), Description(""), DisplayName("Intference Id")>
    Public Property interference_id() As Integer?
        Get
            Return Me._interference_id
        End Get
        Set
            Me._interference_id = Value
        End Set
    End Property
    <Category("CCIpole Int Details"), Description(""), DisplayName("Group Id")>
    Public Property group_id() As Integer?
        Get
            Return Me._group_id
        End Get
        Set
            Me._group_id = Value
        End Set
    End Property
    <Category("CCIpole Int Details"), Description(""), DisplayName("Local Group Id")>
    Public Property local_group_id() As Integer?
        Get
            Return Me._local_group_id
        End Get
        Set
            Me._local_group_id = Value
        End Set
    End Property
    <Category("CCIpole Int Details"), Description(""), DisplayName("Local Intference Id")>
    Public Property local_interference_id() As Integer?
        Get
            Return Me._local_interference_id
        End Get
        Set
            Me._local_interference_id = Value
        End Set
    End Property
    <Category("CCIpole Int Details"), Description(""), DisplayName("Pole Flat")>
    Public Property pole_flat() As Integer?
        Get
            Return Me._pole_flat
        End Get
        Set
            Me._pole_flat = Value
        End Set
    End Property
    <Category("CCIpole Int Details"), Description(""), DisplayName("Horizontal Offset")>
    Public Property horizontal_offset() As Double?
        Get
            Return Me._horizontal_offset
        End Get
        Set
            Me._horizontal_offset = Value
        End Set
    End Property
    <Category("CCIpole Int Details"), Description(""), DisplayName("Rotation")>
    Public Property rotation() As Double?
        Get
            Return Me._rotation
        End Get
        Set
            Me._rotation = Value
        End Set
    End Property
    <Category("CCIpole Int Details"), Description(""), DisplayName("Note")>
    Public Property note() As String
        Get
            Return Me._note
        End Get
        Set
            Me._note = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByRef Parent As PoleIntGroup = Nothing) ', ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(DirectCast(Parent, EDSObject))
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row

        Me.interference_id = DBtoNullableInt(dr.Item("ID"))
        Me.group_id = DBtoNullableInt(dr.Item("group_id"))
        Me.local_group_id = DBtoNullableInt(dr.Item("local_group_id"))
        Me.local_interference_id = DBtoNullableInt(dr.Item("local_interference_id"))
        Me.pole_flat = DBtoNullableInt(dr.Item("pole_flat"))
        Me.horizontal_offset = DBtoNullableDbl(dr.Item("horizontal_offset"))
        Me.rotation = DBtoNullableDbl(dr.Item("rotation"))
        Me.note = DBtoStr(dr.Item("note"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.interference_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID") '(Me.group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_interference_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pole_flat.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.horizontal_offset.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rotation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.note.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("interference_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_interference_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_flat")
        SQLInsertFields = SQLInsertFields.AddtoDBString("horizontal_offset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rotation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("note")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        Dim ParentGroup As PoleIntGroup = TryCast(Me.Parent, PoleIntGroup)

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("interference_id = " & Me.interference_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("group_id = " & ParentGroup.group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_group_id = " & Me.local_group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_interference_id = " & Me.local_interference_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_flat = " & Me.pole_flat.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("horizontal_offset = " & Me.horizontal_offset.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rotation = " & Me.rotation.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("note = " & Me.note.ToString.FormatDBValue)

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
        Dim otherToCompare As PoleIntDetail = TryCast(other, PoleIntDetail)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.interference_id.CheckChange(otherToCompare.interference_id, changes, categoryName, "Intference Id"), Equals, False)
        'Equals = If(Me.group_id.CheckChange(otherToCompare.group_id, changes, categoryName, "Group Id"), Equals, False)
        Equals = If(Me.local_group_id.CheckChange(otherToCompare.local_group_id, changes, categoryName, "Local Group Id"), Equals, False)
        Equals = If(Me.local_interference_id.CheckChange(otherToCompare.local_interference_id, changes, categoryName, "Local Intference Id"), Equals, False)
        Equals = If(Me.pole_flat.CheckChange(otherToCompare.pole_flat, changes, categoryName, "Pole Flat"), Equals, False)
        Equals = If(Me.horizontal_offset.CheckChange(otherToCompare.horizontal_offset, changes, categoryName, "Horizontal Offset"), Equals, False)
        Equals = If(Me.rotation.CheckChange(otherToCompare.rotation, changes, categoryName, "Rotation"), Equals, False)
        Equals = If(Me.note.CheckChange(otherToCompare.note, changes, categoryName, "Note"), Equals, False)


    End Function
#End Region

End Class

Partial Public Class PoleReinfResults
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Results"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.reinforcement_results"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\8 Reinf Result (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Result_INSERT

        SQLInsert = SQLInsert.Replace("[local_section_id]", Me.local_section_id.ToString.FormatDBValue)
        If IsSomething(Me.local_group_id) Then
            SQLInsert = SQLInsert.Replace("[local_group_id]", Me.local_group_id.ToString.FormatDBValue)
        Else
            SQLInsert = SQLInsert.Replace("SELECT @SubLevel2ID = ID FROM pole.reinforcements", "SET @SubLevel2ID = NULL --SELECT @SubLevel2ID = ID FROM pole.reinforcements")
        End If

        SQLInsert = SQLInsert.Replace("[REINF RESULT VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[REINF RESULT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    'Results should only ever be Inserted - MRR
    'Public Overrides Function SQLUpdate() As String
    '    'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\8 Reinf Result (UPDATE).sql")
    '    SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Result_UPDATE

    '    SQLUpdate = SQLUpdate.Replace("[WO]", Me.work_order_seq_num.ToString.FormatDBValue)
    '    SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
    '    SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLUpdate
    'End Function

    'Public Overrides Function SQLDelete() As String
    '    'SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\8 Reinf Result (DELETE).sql")
    '    SQLDelete = CCI_Engineering_Templates.My.Resources.CCIpole_Reinf_Result_DELETE

    '    SQLDelete = SQLDelete.Replace("[WO]", Me.work_order_seq_num.ToString.FormatDBValue)
    '    SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLDelete
    'End Function
#End Region

#Region "Define"
    'Private _ID As Integer?
    Private _result_id As Integer?
    Private _work_order_seq_num As Integer?
    Private _pole_id As Integer?
    Private _section_id As Integer?
    Private _local_section_id As Integer?
    Private _group_id As Integer?
    Private _local_group_id As Integer?
    Private _result_lkup As String
    Private _rating As Double?
    'Private _modified_person_id As Integer?
    'Private _process_stage As String
    'Private _modified_date As DateTime?

    <Category("CCIpole Results"), Description(""), DisplayName("Result Id")>
    Public Property result_id() As Integer?
        Get
            Return Me._result_id
        End Get
        Set
            Me._result_id = Value
        End Set
    End Property
    <Category("CCIpole Results"), Description(""), DisplayName("Work Order Seq Num")>
    Public Property work_order_seq_num() As Integer?
        Get
            Return Me._work_order_seq_num
        End Get
        Set
            Me._work_order_seq_num = Value
        End Set
    End Property
    <Category("CCIpole Results"), Description(""), DisplayName("Pole Id")>
    Public Property pole_id() As Integer?
        Get
            Return Me._pole_id
        End Get
        Set
            Me._pole_id = Value
        End Set
    End Property
    <Category("CCIpole Results"), Description(""), DisplayName("Section Id")>
    Public Property section_id() As Integer?
        Get
            Return Me._section_id
        End Get
        Set
            Me._section_id = Value
        End Set
    End Property

    <Category("CCIpole Results"), Description(""), DisplayName("Local Section Id")>
    Public Property local_section_id() As Integer?
        Get
            Return Me._local_section_id
        End Get
        Set
            Me._local_section_id = Value
        End Set
    End Property
    <Category("CCIpole Results"), Description(""), DisplayName("Group Id")>
    Public Property group_id() As Integer?
        Get
            Return Me._group_id
        End Get
        Set
            Me._group_id = Value
        End Set
    End Property
    <Category("CCIpole Results"), Description(""), DisplayName("Local Group Id")>
    Public Property local_group_id() As Integer?
        Get
            Return Me._local_group_id
        End Get
        Set
            Me._local_group_id = Value
        End Set
    End Property
    <Category("CCIpole Results"), Description(""), DisplayName("Result Lkup")>
    Public Property result_lkup() As String
        Get
            Return Me._result_lkup
        End Get
        Set
            Me._result_lkup = Value
        End Set
    End Property
    <Category("CCIpole Results"), Description(""), DisplayName("Rating")>
    Public Property rating() As Double?
        Get
            Return Me._rating
        End Get
        Set
            Me._rating = Value
        End Set
    End Property
    '<Category("CCIpole Results"), Description(""), DisplayName("Modified Person Id")>
    'Public Property modified_person_id() As Integer?
    '    Get
    '        Return Me._modified_person_id
    '    End Get
    '    Set
    '        Me._modified_person_id = Value
    '    End Set
    'End Property
    '<Category("CCIpole Results"), Description(""), DisplayName("Process Stage")>
    'Public Property process_stage() As String
    '    Get
    '        Return Me._process_stage
    '    End Get
    '    Set
    '        Me._process_stage = Value
    '    End Set
    'End Property
    '<Category("CCIpole Results"), Description(""), DisplayName("Modified Date")>
    'Public Property modified_date() As DateTime?
    '    Get
    '        Return Me._modified_date
    '    End Get
    '    Set
    '        Me._modified_date = Value
    '    End Set
    'End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByVal Parent As EDSObject = Nothing) ', ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row

        Me.result_id = DBtoNullableInt(dr.Item("ID"))
        Me.work_order_seq_num = DBtoNullableInt(dr.Item("work_order_seq_num")) 'Change to Parent WO downstream - MRR
        Me.pole_id = DBtoNullableInt(dr.Item("pole_id"))
        Me.section_id = DBtoNullableInt(dr.Item("section_id"))
        Me.local_section_id = DBtoNullableInt(dr.Item("local_section_id"))
        Me.group_id = DBtoNullableInt(dr.Item("group_id"))
        Me.local_group_id = DBtoNullableInt(dr.Item("local_group_id"))
        Me.result_lkup = DBtoStr(dr.Item("result_lkup"))
        Me.rating = DBtoNullableDbl(dr.Item("rating"))
        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
        'Me.modified_date = DBtoStr(dr.Item("modified_date"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Parent.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") '(Me.pole_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID") '(Me.section_id.ToString.FormatDBValue) 
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_section_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID") '(Me.group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_group_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("result_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pole_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_section_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_group_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        Dim ParentPole As Pole = TryCast(Me.Parent, Pole)

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_id = " & Me.result_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Parent.work_order_seq_num.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_id = " & ParentPole.pole_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("section_id = " & Me.section_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_section_id = " & Me.local_section_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("group_id = " & Me.group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_group_id = " & Me.local_group_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

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
        Dim otherToCompare As PoleReinfResults = TryCast(other, PoleReinfResults)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.result_id.CheckChange(otherToCompare.result_id, changes, categoryName, "Result Id"), Equals, False)
        Equals = If(Me.work_order_seq_num.CheckChange(otherToCompare.work_order_seq_num, changes, categoryName, "Work Order Seq Num"), Equals, False)
        Equals = If(Me.pole_id.CheckChange(otherToCompare.pole_id, changes, categoryName, "Pole Id"), Equals, False)
        'Equals = If(Me.section_id.CheckChange(otherToCompare.section_id, changes, categoryName, "Section Id"), Equals, False)
        Equals = If(Me.local_section_id.CheckChange(otherToCompare.local_section_id, changes, categoryName, "Local Section Id"), Equals, False)
        'Equals = If(Me.group_id.CheckChange(otherToCompare.group_id, changes, categoryName, "Group Id"), Equals, False)
        Equals = If(Me.local_group_id.CheckChange(otherToCompare.local_group_id, changes, categoryName, "Local Group Id"), Equals, False)
        Equals = If(Me.result_lkup.CheckChange(otherToCompare.result_lkup, changes, categoryName, "Result Lkup"), Equals, False)
        Equals = If(Me.rating.CheckChange(otherToCompare.rating, changes, categoryName, "Rating"), Equals, False)
        'Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        'Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)
        'Equals = If(Me.modified_date.CheckChange(otherToCompare.modified_date, changes, categoryName, "Modified Date"), Equals, False)

    End Function
#End Region

End Class

Partial Public Class PoleMatlProp
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Custom Matls"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.pole_matls"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\DB Matl (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_DB_Matl_INSERT

        SQLInsert = SQLInsert.Replace("[MATL DB FIELDS AND VALUES]", Me.SQLUpdateFieldsandValues.Replace(", ", " AND "))
        SQLInsert = SQLInsert.Replace("= NULL ", "IS NULL ")
        SQLInsert = SQLInsert.Replace("[MATL DB VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[MATL DB FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    'Public Overrides Function SQLUpdate() As String 'Probably do not even need UPDATE or DELETE functionality. Should only be adding items to the reference DBs - MRR
    '    SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\Prop Matl (UPDATE).sql")
    '    SQLUpdate = SQLUpdate.Replace("[ID]", Me.pole_id.ToString.FormatDBValue)
    '    SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
    '    SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLUpdate
    'End Function

    'Public Overrides Function SQLDelete() As String
    '    SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\Prop Matl (DELETE).sql")
    '    SQLDelete = SQLDelete.Replace("[ID]", Me.pole_id.ToString.FormatDBValue)
    '    SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLDelete
    'End Function
#End Region

#Region "Define"
    Private _matl_id As Integer?
    Private _local_matl_id As Integer?
    Private _name As String
    Private _fy As Double?
    Private _fu As Double?
    Private _ind_default As Boolean?

    <Category("CCIpole Matl DB"), Description(""), DisplayName("Matl Id")>
    Public Property matl_id() As Integer?
        Get
            Return Me._matl_id
        End Get
        Set
            Me._matl_id = Value
        End Set
    End Property
    <Category("CCIpole Matl DB"), Description(""), DisplayName("Local Matl Id")>
    Public Property local_matl_id() As Integer?
        Get
            Return Me._local_matl_id
        End Get
        Set
            Me._local_matl_id = Value
        End Set
    End Property
    <Category("CCIpole Matl DB"), Description(""), DisplayName("Name")>
    Public Property name() As String
        Get
            Return Me._name
        End Get
        Set
            Me._name = Value
        End Set
    End Property
    <Category("CCIpole Matl DB"), Description(""), DisplayName("Fy")>
    Public Property fy() As Double?
        Get
            Return Me._fy
        End Get
        Set
            Me._fy = Value
        End Set
    End Property
    <Category("CCIpole Matl DB"), Description(""), DisplayName("Fu")>
    Public Property fu() As Double?
        Get
            Return Me._fu
        End Get
        Set
            Me._fu = Value
        End Set
    End Property
    <Category("CCIpole Matl DB"), Description(""), DisplayName("Ind Default")>
    Public Property ind_default() As Boolean?
        Get
            Return Me._ind_default
        End Get
        Set
            Me._ind_default = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByVal Parent As EDSObject = Nothing) ', ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row

        Me.matl_id = DBtoNullableInt(dr.Item("ID"))
        Me.local_matl_id = DBtoNullableInt(dr.Item("local_matl_id"))
        Me.name = DBtoStr(dr.Item("name"))
        Me.fy = DBtoNullableDbl(dr.Item("fy"))
        Me.fu = DBtoNullableDbl(dr.Item("fu"))
        Me.ind_default = DBtoNullableBool(dr.Item("ind_default"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.matl_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") '(Me.pole_id.ToString.FormatDBValue) - Does pole_id need to be added? - MRR
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_matl_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.name.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ind_default.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("matl_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("pole_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_matl_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("name")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fy")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fu")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ind_default")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("matl_id = " & Me.matl_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_id = " & Me.pole_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_matl_id = " & Me.local_matl_id.ToString.FormatDBValue) 'Commented out so when searching if a matching entry exists within DB already, the local ID wont disqualify it
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("name = " & Me.name.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("fy = " & Me.fy.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("fu = " & Me.fu.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ind_default = " & Me.ind_default.ToString.FormatDBValue)

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
        Dim otherToCompare As PoleMatlProp = TryCast(other, PoleMatlProp)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.matl_id.CheckChange(otherToCompare.matl_id, changes, categoryName, "Matl Id"), Equals, False)
        'Equals = If(Me.pole_id.CheckChange(otherToCompare.pole_id, changes, categoryName, "Pole Id"), Equals, False)
        Equals = If(Me.local_matl_id.CheckChange(otherToCompare.local_matl_id, changes, categoryName, "Local Matl Id"), Equals, False)
        Equals = If(Me.name.CheckChange(otherToCompare.name, changes, categoryName, "Name"), Equals, False)
        Equals = If(Me.fy.CheckChange(otherToCompare.fy, changes, categoryName, "Fy"), Equals, False)
        Equals = If(Me.fu.CheckChange(otherToCompare.fu, changes, categoryName, "Fu"), Equals, False)
        Equals = If(Me.ind_default.CheckChange(otherToCompare.ind_default, changes, categoryName, "Ind Default"), Equals, False)

    End Function
#End Region

End Class

Partial Public Class PoleBoltProp
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Custom Bolts"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.pole_bolts"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\DB Bolt (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_DB_Bolt_INSERT

        SQLInsert = SQLInsert.Replace("[BOLT DB FIELDS AND VALUES]", Me.SQLUpdateFieldsandValues.Replace(", ", " AND "))
        SQLInsert = SQLInsert.Replace("= NULL ", "IS NULL ")
        SQLInsert = SQLInsert.Replace("[BOLT DB VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[BOLT DB FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    'Public Overrides Function SQLUpdate() As String
    '    SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\Prop Bolt (UPDATE).sql")
    '    SQLUpdate = SQLUpdate.Replace("[ID]", Me.pole_id.ToString.FormatDBValue)
    '    SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
    '    SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLUpdate
    'End Function

    'Public Overrides Function SQLDelete() As String
    '    SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\Prop Bolt (DELETE).sql")
    '    SQLDelete = SQLDelete.Replace("[ID]", Me.pole_id.ToString.FormatDBValue)
    '    SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLDelete
    'End Function
#End Region

#Region "Define"
    Private _bolt_id As Integer?
    Private _local_bolt_id As Integer?
    Private _name As String
    Private _description As String
    Private _diam As Double?
    Private _area As Double?
    Private _fu_bolt As Double?
    Private _sleeve_diam_out As Double?
    Private _sleeve_diam_in As Double?
    Private _fu_sleeve As Double?
    Private _bolt_n_sleeve_shear_revF As Double?
    Private _bolt_x_sleeve_shear_revF As Double?
    Private _bolt_n_sleeve_shear_revG As Double?
    Private _bolt_x_sleeve_shear_revG As Double?
    Private _bolt_n_sleeve_shear_revH As Double?
    Private _bolt_x_sleeve_shear_revH As Double?
    Private _rb_applied_revH As Boolean?
    Private _ind_default As Boolean?

    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Bolt Id")>
    Public Property bolt_id() As Integer?
        Get
            Return Me._bolt_id
        End Get
        Set
            Me._bolt_id = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Local Bolt Id")>
    Public Property local_bolt_id() As Integer?
        Get
            Return Me._local_bolt_id
        End Get
        Set
            Me._local_bolt_id = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Name")>
    Public Property name() As String
        Get
            Return Me._name
        End Get
        Set
            Me._name = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Description")>
    Public Property description() As String
        Get
            Return Me._description
        End Get
        Set
            Me._description = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Diam")>
    Public Property diam() As Double?
        Get
            Return Me._diam
        End Get
        Set
            Me._diam = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Area")>
    Public Property area() As Double?
        Get
            Return Me._area
        End Get
        Set
            Me._area = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Fu Bolt")>
    Public Property fu_bolt() As Double?
        Get
            Return Me._fu_bolt
        End Get
        Set
            Me._fu_bolt = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Sleeve Diam Out")>
    Public Property sleeve_diam_out() As Double?
        Get
            Return Me._sleeve_diam_out
        End Get
        Set
            Me._sleeve_diam_out = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Sleeve Diam In")>
    Public Property sleeve_diam_in() As Double?
        Get
            Return Me._sleeve_diam_in
        End Get
        Set
            Me._sleeve_diam_in = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Fu Sleeve")>
    Public Property fu_sleeve() As Double?
        Get
            Return Me._fu_sleeve
        End Get
        Set
            Me._fu_sleeve = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Bolt N Sleeve Shear Revf")>
    Public Property bolt_n_sleeve_shear_revF() As Double?
        Get
            Return Me._bolt_n_sleeve_shear_revF
        End Get
        Set
            Me._bolt_n_sleeve_shear_revF = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Bolt X Sleeve Shear Revf")>
    Public Property bolt_x_sleeve_shear_revF() As Double?
        Get
            Return Me._bolt_x_sleeve_shear_revF
        End Get
        Set
            Me._bolt_x_sleeve_shear_revF = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Bolt N Sleeve Shear Revg")>
    Public Property bolt_n_sleeve_shear_revG() As Double?
        Get
            Return Me._bolt_n_sleeve_shear_revG
        End Get
        Set
            Me._bolt_n_sleeve_shear_revG = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Bolt X Sleeve Shear Revg")>
    Public Property bolt_x_sleeve_shear_revG() As Double?
        Get
            Return Me._bolt_x_sleeve_shear_revG
        End Get
        Set
            Me._bolt_x_sleeve_shear_revG = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Bolt N Sleeve Shear Revh")>
    Public Property bolt_n_sleeve_shear_revH() As Double?
        Get
            Return Me._bolt_n_sleeve_shear_revH
        End Get
        Set
            Me._bolt_n_sleeve_shear_revH = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Bolt X Sleeve Shear Revh")>
    Public Property bolt_x_sleeve_shear_revH() As Double?
        Get
            Return Me._bolt_x_sleeve_shear_revH
        End Get
        Set
            Me._bolt_x_sleeve_shear_revH = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Rb Applied Revh")>
    Public Property rb_applied_revH() As Boolean?
        Get
            Return Me._rb_applied_revH
        End Get
        Set
            Me._rb_applied_revH = Value
        End Set
    End Property
    <Category("CCIpole Bolt DB"), Description(""), DisplayName("Ind Default")>
    Public Property ind_default() As Boolean?
        Get
            Return Me._ind_default
        End Get
        Set
            Me._ind_default = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByVal Parent As EDSObject = Nothing) ', ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row

        Me.bolt_id = DBtoNullableInt(dr.Item("ID"))
        Me.local_bolt_id = DBtoNullableInt(dr.Item("local_bolt_id"))
        Me.name = DBtoStr(dr.Item("name"))
        Me.description = DBtoStr(dr.Item("description"))
        Me.diam = DBtoNullableDbl(dr.Item("diam"))
        Me.area = DBtoNullableDbl(dr.Item("area"))
        Me.fu_bolt = DBtoNullableDbl(dr.Item("fu_bolt"))
        Me.sleeve_diam_out = DBtoNullableDbl(dr.Item("sleeve_diam_out"))
        Me.sleeve_diam_in = DBtoNullableDbl(dr.Item("sleeve_diam_in"))
        Me.fu_sleeve = DBtoNullableDbl(dr.Item("fu_sleeve"))
        Me.bolt_n_sleeve_shear_revF = DBtoNullableDbl(dr.Item("bolt_n_sleeve_shear_revF"))
        Me.bolt_x_sleeve_shear_revF = DBtoNullableDbl(dr.Item("bolt_x_sleeve_shear_revF"))
        Me.bolt_n_sleeve_shear_revG = DBtoNullableDbl(dr.Item("bolt_n_sleeve_shear_revG"))
        Me.bolt_x_sleeve_shear_revG = DBtoNullableDbl(dr.Item("bolt_x_sleeve_shear_revG"))
        Me.bolt_n_sleeve_shear_revH = DBtoNullableDbl(dr.Item("bolt_n_sleeve_shear_revH"))
        Me.bolt_x_sleeve_shear_revH = DBtoNullableDbl(dr.Item("bolt_x_sleeve_shear_revH"))
        Me.rb_applied_revH = DBtoNullableBool(dr.Item("rb_applied_revH"))
        Me.ind_default = DBtoNullableBool(dr.Item("ind_default"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") '(Me.pole_id.ToString.FormatDBValue) - Does pole_id need to be added? - MRR
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_bolt_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.name.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.description.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.diam.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.area.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_bolt.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.sleeve_diam_out.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.sleeve_diam_in.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_sleeve.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_n_sleeve_shear_revF.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_x_sleeve_shear_revF.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_n_sleeve_shear_revG.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_x_sleeve_shear_revG.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_n_sleeve_shear_revH.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_x_sleeve_shear_revH.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rb_applied_revH.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ind_default.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("pole_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_bolt_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("name")
        SQLInsertFields = SQLInsertFields.AddtoDBString("description")
        SQLInsertFields = SQLInsertFields.AddtoDBString("diam")
        SQLInsertFields = SQLInsertFields.AddtoDBString("area")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_bolt")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sleeve_diam_out")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sleeve_diam_in")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_sleeve")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_n_sleeve_shear_revF")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_x_sleeve_shear_revF")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_n_sleeve_shear_revG")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_x_sleeve_shear_revG")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_n_sleeve_shear_revH")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_x_sleeve_shear_revH")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rb_applied_revH")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ind_default")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_id = " & Me.bolt_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_id = " & Me.pole_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_bolt_id = " & Me.local_bolt_id.ToString.FormatDBValue) 'Commented out so when searching if a matching entry exists within DB already, the local ID wont disqualify it
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("name = " & Me.name.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("description = " & Me.description.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("diam = " & Me.diam.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("area = " & Me.area.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("fu_bolt = " & Me.fu_bolt.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sleeve_diam_out = " & Me.sleeve_diam_out.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sleeve_diam_in = " & Me.sleeve_diam_in.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("fu_sleeve = " & Me.fu_sleeve.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_n_sleeve_shear_revF = " & Me.bolt_n_sleeve_shear_revF.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_x_sleeve_shear_revF = " & Me.bolt_x_sleeve_shear_revF.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_n_sleeve_shear_revG = " & Me.bolt_n_sleeve_shear_revG.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_x_sleeve_shear_revG = " & Me.bolt_x_sleeve_shear_revG.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_n_sleeve_shear_revH = " & Me.bolt_n_sleeve_shear_revH.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_x_sleeve_shear_revH = " & Me.bolt_x_sleeve_shear_revH.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rb_applied_revH = " & Me.rb_applied_revH.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ind_default = " & Me.ind_default.ToString.FormatDBValue)


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
        Dim otherToCompare As PoleBoltProp = TryCast(other, PoleBoltProp)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.bolt_id.CheckChange(otherToCompare.bolt_id, changes, categoryName, "Bolt Id"), Equals, False)
        'Equals = If(Me.pole_id.CheckChange(otherToCompare.pole_id, changes, categoryName, "Pole Id"), Equals, False)
        Equals = If(Me.local_bolt_id.CheckChange(otherToCompare.local_bolt_id, changes, categoryName, "Local Bolt Id"), Equals, False)
        Equals = If(Me.name.CheckChange(otherToCompare.name, changes, categoryName, "Name"), Equals, False)
        Equals = If(Me.description.CheckChange(otherToCompare.description, changes, categoryName, "Description"), Equals, False)
        Equals = If(Me.diam.CheckChange(otherToCompare.diam, changes, categoryName, "Diam"), Equals, False)
        Equals = If(Me.area.CheckChange(otherToCompare.area, changes, categoryName, "Area"), Equals, False)
        Equals = If(Me.fu_bolt.CheckChange(otherToCompare.fu_bolt, changes, categoryName, "Fu Bolt"), Equals, False)
        Equals = If(Me.sleeve_diam_out.CheckChange(otherToCompare.sleeve_diam_out, changes, categoryName, "Sleeve Diam Out"), Equals, False)
        Equals = If(Me.sleeve_diam_in.CheckChange(otherToCompare.sleeve_diam_in, changes, categoryName, "Sleeve Diam In"), Equals, False)
        Equals = If(Me.fu_sleeve.CheckChange(otherToCompare.fu_sleeve, changes, categoryName, "Fu Sleeve"), Equals, False)
        Equals = If(Me.bolt_n_sleeve_shear_revF.CheckChange(otherToCompare.bolt_n_sleeve_shear_revF, changes, categoryName, "Bolt N Sleeve Shear Revf"), Equals, False)
        Equals = If(Me.bolt_x_sleeve_shear_revF.CheckChange(otherToCompare.bolt_x_sleeve_shear_revF, changes, categoryName, "Bolt X Sleeve Shear Revf"), Equals, False)
        Equals = If(Me.bolt_n_sleeve_shear_revG.CheckChange(otherToCompare.bolt_n_sleeve_shear_revG, changes, categoryName, "Bolt N Sleeve Shear Revg"), Equals, False)
        Equals = If(Me.bolt_x_sleeve_shear_revG.CheckChange(otherToCompare.bolt_x_sleeve_shear_revG, changes, categoryName, "Bolt X Sleeve Shear Revg"), Equals, False)
        Equals = If(Me.bolt_n_sleeve_shear_revH.CheckChange(otherToCompare.bolt_n_sleeve_shear_revH, changes, categoryName, "Bolt N Sleeve Shear Revh"), Equals, False)
        Equals = If(Me.bolt_x_sleeve_shear_revH.CheckChange(otherToCompare.bolt_x_sleeve_shear_revH, changes, categoryName, "Bolt X Sleeve Shear Revh"), Equals, False)
        Equals = If(Me.rb_applied_revH.CheckChange(otherToCompare.rb_applied_revH, changes, categoryName, "Rb Applied Revh"), Equals, False)
        Equals = If(Me.ind_default.CheckChange(otherToCompare.ind_default, changes, categoryName, "Ind Default"), Equals, False)


    End Function
#End Region

End Class

Partial Public Class PoleReinfProp
    Inherits EDSObjectWithQueries

#Region "Inherited"
    Public Overrides ReadOnly Property EDSObjectName As String = "Pole Custom Reinfs"
    Public Overrides ReadOnly Property EDSTableName As String = "pole.pole_reinforcements"

    Public Overrides Function SQLInsert() As String
        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIpole\DB Reinf (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIpole_DB_Reinf_INSERT

        SQLInsert = SQLInsert.Replace("[REINF DB FIELDS AND VALUES]", Me.SQLUpdateFieldsandValues.Replace(", ", " AND "))
        SQLInsert = SQLInsert.Replace("= NULL ", "IS NULL ")
        SQLInsert = SQLInsert.Replace("[REINF DB VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[REINF DB FIELDS]", Me.SQLInsertFields)

        Dim ParentPole As Pole = TryCast(Me.Parent, Pole)

        'Matl
        If IsSomething(Me.matl_id) And Me.matl_id <> 0 Then
            SQLInsert = SQLInsert.Replace("[MATL ID]", Me.matl_id.ToString.FormatDBValue)
        Else
            SQLInsert = SQLInsert.Replace("[MATL ID]", "NULL")
            If IsSomething(ParentPole) Then
                For Each dbrow As PoleMatlProp In ParentPole.matls
                    If Me.local_matl_id = dbrow.local_matl_id Then 'And Me.local_matl_id > 17 Then 'Matching, Non-Standard Materials
                        SQLInsert = SQLInsert.Replace("--[MATL DB SUBQUERY]", dbrow.SQLInsert)
                    End If
                Next
            End If
        End If

        'Top Bolt
        If IsSomething(Me.bolt_id_top) And Me.bolt_id_top <> 0 Then
            SQLInsert = SQLInsert.Replace("[TOP BOLT ID]", Me.bolt_id_top.ToString.FormatDBValue)
        Else
            SQLInsert = SQLInsert.Replace("[TOP BOLT ID]", "NULL")
            If IsSomething(ParentPole) Then
                For Each dbrow As PoleBoltProp In ParentPole.bolts
                    If Me.local_bolt_id_top = dbrow.local_bolt_id Then '
                        SQLInsert = SQLInsert.Replace("--[TOP BOLT DB SUBQUERY]", dbrow.SQLInsert)
                        SQLInsert = SQLInsert.Replace("@BoltID", "@TopBoltID")
                    End If
                Next
            End If
            If Me.local_bolt_id_top = 0 Then
                SQLInsert = SQLInsert.Replace("bolt_id_top = @TopBoltID", "bolt_id_top IS NULL")
            End If
        End If

        'Bot Bolt
        If IsSomething(Me.bolt_id_bot) And Me.bolt_id_bot <> 0 Then
            SQLInsert = SQLInsert.Replace("[BOT BOLT ID]", Me.bolt_id_bot.ToString.FormatDBValue)
        Else
            SQLInsert = SQLInsert.Replace("[BOT BOLT ID]", "NULL")
            If IsSomething(ParentPole) Then
                For Each dbrow As PoleBoltProp In ParentPole.bolts
                    If Me.local_bolt_id_bot = dbrow.local_bolt_id Then '
                        SQLInsert = SQLInsert.Replace("--[BOT BOLT DB SUBQUERY]", dbrow.SQLInsert)
                        SQLInsert = SQLInsert.Replace("@BoltID", "@BotBoltID")
                    End If
                Next
            End If
            If Me.local_bolt_id_bot = 0 Then
                SQLInsert = SQLInsert.Replace("bolt_id_bot = @BotBoltID", "bolt_id_bot IS NULL")
            End If
        End If

        'If IsSomething(Me.reinf_id) And Me.reinf_id <> 0 And Me.reinf_id <= 10 Then 'Default Reinf Used - no need to pull in subqueries in order to create Reinf Group
        '    SQLInsert = SQLInsert.Replace("--[REINF DB INSTERT]",
        '                                  "SET @SubLevel4ID = 6" & vbNewLine &
        '                                  "SET @TopBoltID = 1" & vbNewLine &
        '                                  "SET @BotBoltID = 1" & vbNewLine &
        '                                  "SET @SubLevel2ID = " & Me.reinf_id.ToString.FormatDBValue)
        'ElseIf IsSomething(Me.reinf_id) And Me.reinf_id <> 0 And Me.reinf_id > 10 And Me.reinf_id < 25 Then
        '    SQLInsert = SQLInsert.Replace("--[REINF DB INSTERT]",
        '                                  "SET @SubLevel4ID = 12" & vbNewLine &
        '                                  "SET @TopBoltID = 1" & vbNewLine &
        '                                  "SET @BotBoltID = 1" & vbNewLine &
        '                                  "SET @SubLevel2ID = " & Me.reinf_id.ToString.FormatDBValue)
        'ElseIf IsSomething(Me.reinf_id) And Me.reinf_id <> 0 And Me.reinf_id >= 25 And Me.reinf_id <= 38 Then
        '    SQLInsert = SQLInsert.Replace("--[REINF DB INSTERT]",
        '                                  "SET @SubLevel4ID = 12" & vbNewLine &
        '                                  "SET @TopBoltID = 1" & vbNewLine &
        '                                  "SET @BotBoltID = NULL" & vbNewLine &
        '                                  "SET @SubLevel2ID = " & Me.reinf_id.ToString.FormatDBValue)
        'ElseIf IsSomething(Me.reinf_id) And Me.reinf_id <> 0 And Me.reinf_id > 38 And Me.reinf_id < 58 Then
        '    SQLInsert = SQLInsert.Replace("--[REINF DB INSTERT]",
        '                                  "SET @SubLevel4ID = 12" & vbNewLine &
        '                                  "SET @TopBoltID = 1" & vbNewLine &
        '                                  "SET @BotBoltID = 1" & vbNewLine &
        '                                  "SET @SubLevel2ID = " & Me.reinf_id.ToString.FormatDBValue)
        'End If


        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert
    End Function

    'Public Overrides Function SQLUpdate() As String
    '    SQLUpdate = QueryBuilderFromFile(queryPath & "CCIpole\Prop Reinf (UPDATE).sql")
    '    SQLUpdate = SQLUpdate.Replace("[ID]", Me.pole_id.ToString.FormatDBValue)
    '    SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
    '    SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLUpdate
    'End Function

    'Public Overrides Function SQLDelete() As String
    '    SQLDelete = QueryBuilderFromFile(queryPath & "CCIpole\Prop Reinf (DELETE).sql")
    '    SQLDelete = SQLDelete.Replace("[ID]", Me.pole_id.ToString.FormatDBValue)
    '    SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLDelete
    'End Function
#End Region

#Region "Define"
    Private _reinf_id As Integer?
    Private _local_reinf_id As Integer?                     'Local ID
    Private _name As String
    Private _type As String
    Private _b As Double?
    Private _h As Double?
    Private _sr_diam As Double?
    Private _channel_thkns_web As Double?
    Private _channel_thkns_flange As Double?
    Private _channel_eo As Double?
    Private _channel_J As Double?
    Private _channel_Cw As Double?
    Private _area_gross As Double?
    Private _centroid As Double?
    Private _istension As Boolean?
    Private _matl_id As Integer?
    Private _local_matl_id As Integer?                      'Local ID
    Private _Ix As Double?
    Private _Iy As Double?
    Private _Lu As Double?
    Private _Kx As Double?
    Private _Ky As Double?
    Private _bolt_hole_size As Double?
    Private _area_net As Double?
    Private _shear_lag As Double?
    Private _connection_type_bot As String
    Private _connection_cap_revF_bot As Double?
    Private _connection_cap_revG_bot As Double?
    Private _connection_cap_revH_bot As Double?
    Private _bolt_id_bot As Integer?
    Private _local_bolt_id_bot As Integer?                  'Local ID
    Private _bolt_N_or_X_bot As String
    Private _bolt_num_bot As Integer?
    Private _bolt_spacing_bot As Double?
    Private _bolt_edge_dist_bot As Double?
    Private _FlangeOrBP_connected_bot As Boolean?
    Private _weld_grade_bot As Double?
    Private _weld_trans_type_bot As String
    Private _weld_trans_length_bot As Double?
    Private _weld_groove_depth_bot As Double?
    Private _weld_groove_angle_bot As Integer?
    Private _weld_trans_fillet_size_bot As Double?
    Private _weld_trans_eff_throat_bot As Double?
    Private _weld_long_type_bot As String
    Private _weld_long_length_bot As Double?
    Private _weld_long_fillet_size_bot As Double?
    Private _weld_long_eff_throat_bot As Double?
    Private _top_bot_connections_symmetrical As Boolean?
    Private _connection_type_top As String
    Private _connection_cap_revF_top As Double?
    Private _connection_cap_revG_top As Double?
    Private _connection_cap_revH_top As Double?
    Private _bolt_id_top As Integer?
    Private _local_bolt_id_top As Integer?                  'Local ID
    Private _bolt_N_or_X_top As String
    Private _bolt_num_top As Integer?
    Private _bolt_spacing_top As Double?
    Private _bolt_edge_dist_top As Double?
    Private _FlangeOrBP_connected_top As Boolean?
    Private _weld_grade_top As Double?
    Private _weld_trans_type_top As String
    Private _weld_trans_length_top As Double?
    Private _weld_groove_depth_top As Double?
    Private _weld_groove_angle_top As Integer?
    Private _weld_trans_fillet_size_top As Double?
    Private _weld_trans_eff_throat_top As Double?
    Private _weld_long_type_top As String
    Private _weld_long_length_top As Double?
    Private _weld_long_fillet_size_top As Double?
    Private _weld_long_eff_throat_top As Double?
    Private _conn_length_channel As Double?
    Private _conn_length_bot As Double?
    Private _conn_length_top As Double?
    Private _cap_comp_xx_f As Double?
    Private _cap_comp_yy_f As Double?
    Private _cap_tens_yield_f As Double?
    Private _cap_tens_rupture_f As Double?
    Private _cap_shear_f As Double?
    Private _cap_bolt_shear_bot_f As Double?
    Private _cap_bolt_shear_top_f As Double?
    Private _cap_boltshaft_bearing_nodeform_bot_f As Double?
    Private _cap_boltshaft_bearing_deform_bot_f As Double?
    Private _cap_boltshaft_bearing_nodeform_top_f As Double?
    Private _cap_boltshaft_bearing_deform_top_f As Double?
    Private _cap_boltreinf_bearing_nodeform_bot_f As Double?
    Private _cap_boltreinf_bearing_deform_bot_f As Double?
    Private _cap_boltreinf_bearing_nodeform_top_f As Double?
    Private _cap_boltreinf_bearing_deform_top_f As Double?
    Private _cap_weld_trans_bot_f As Double?
    Private _cap_weld_long_bot_f As Double?
    Private _cap_weld_trans_top_f As Double?
    Private _cap_weld_long_top_f As Double?
    Private _cap_comp_xx_g As Double?
    Private _cap_comp_yy_g As Double?
    Private _cap_tens_yield_g As Double?
    Private _cap_tens_rupture_g As Double?
    Private _cap_shear_g As Double?
    Private _cap_bolt_shear_bot_g As Double?
    Private _cap_bolt_shear_top_g As Double?
    Private _cap_boltshaft_bearing_nodeform_bot_g As Double?
    Private _cap_boltshaft_bearing_deform_bot_g As Double?
    Private _cap_boltshaft_bearing_nodeform_top_g As Double?
    Private _cap_boltshaft_bearing_deform_top_g As Double?
    Private _cap_boltreinf_bearing_nodeform_bot_g As Double?
    Private _cap_boltreinf_bearing_deform_bot_g As Double?
    Private _cap_boltreinf_bearing_nodeform_top_g As Double?
    Private _cap_boltreinf_bearing_deform_top_g As Double?
    Private _cap_weld_trans_bot_g As Double?
    Private _cap_weld_long_bot_g As Double?
    Private _cap_weld_trans_top_g As Double?
    Private _cap_weld_long_top_g As Double?
    Private _cap_comp_xx_h As Double?
    Private _cap_comp_yy_h As Double?
    Private _cap_tens_yield_h As Double?
    Private _cap_tens_rupture_h As Double?
    Private _cap_shear_h As Double?
    Private _cap_bolt_shear_bot_h As Double?
    Private _cap_bolt_shear_top_h As Double?
    Private _cap_boltshaft_bearing_nodeform_bot_h As Double?
    Private _cap_boltshaft_bearing_deform_bot_h As Double?
    Private _cap_boltshaft_bearing_nodeform_top_h As Double?
    Private _cap_boltshaft_bearing_deform_top_h As Double?
    Private _cap_boltreinf_bearing_nodeform_bot_h As Double?
    Private _cap_boltreinf_bearing_deform_bot_h As Double?
    Private _cap_boltreinf_bearing_nodeform_top_h As Double?
    Private _cap_boltreinf_bearing_deform_top_h As Double?
    Private _cap_weld_trans_bot_h As Double?
    Private _cap_weld_long_bot_h As Double?
    Private _cap_weld_trans_top_h As Double?
    Private _cap_weld_long_top_h As Double?
    Private _ind_default As Boolean?
    'Public Property matls As New List(Of PoleMatlProp)
    'Public Property bolts As New List(Of PoleBoltProp)

    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Reinf Id")>
    Public Property reinf_id() As Integer?
        Get
            Return Me._reinf_id
        End Get
        Set
            Me._reinf_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Local Reinf Id")>
    Public Property local_reinf_id() As Integer?
        Get
            Return Me._local_reinf_id
        End Get
        Set
            Me._local_reinf_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Name")>
    Public Property name() As String
        Get
            Return Me._name
        End Get
        Set
            Me._name = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Type")>
    Public Property type() As String
        Get
            Return Me._type
        End Get
        Set
            Me._type = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("B")>
    Public Property b() As Double?
        Get
            Return Me._b
        End Get
        Set
            Me._b = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("H")>
    Public Property h() As Double?
        Get
            Return Me._h
        End Get
        Set
            Me._h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Sr Diam")>
    Public Property sr_diam() As Double?
        Get
            Return Me._sr_diam
        End Get
        Set
            Me._sr_diam = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Channel Thkns Web")>
    Public Property channel_thkns_web() As Double?
        Get
            Return Me._channel_thkns_web
        End Get
        Set
            Me._channel_thkns_web = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Channel Thkns Flange")>
    Public Property channel_thkns_flange() As Double?
        Get
            Return Me._channel_thkns_flange
        End Get
        Set
            Me._channel_thkns_flange = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Channel Eo")>
    Public Property channel_eo() As Double?
        Get
            Return Me._channel_eo
        End Get
        Set
            Me._channel_eo = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Channel J")>
    Public Property channel_J() As Double?
        Get
            Return Me._channel_J
        End Get
        Set
            Me._channel_J = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Channel Cw")>
    Public Property channel_Cw() As Double?
        Get
            Return Me._channel_Cw
        End Get
        Set
            Me._channel_Cw = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Area Gross")>
    Public Property area_gross() As Double?
        Get
            Return Me._area_gross
        End Get
        Set
            Me._area_gross = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Centroid")>
    Public Property centroid() As Double?
        Get
            Return Me._centroid
        End Get
        Set
            Me._centroid = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Istension")>
    Public Property istension() As Boolean?
        Get
            Return Me._istension
        End Get
        Set
            Me._istension = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Matl Id")>
    Public Property matl_id() As Integer?
        Get
            Return Me._matl_id
        End Get
        Set
            Me._matl_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Local Matl Id")>
    Public Property local_matl_id() As Integer?
        Get
            Return Me._local_matl_id
        End Get
        Set
            Me._local_matl_id = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Ix")>
    Public Property Ix() As Double?
        Get
            Return Me._Ix
        End Get
        Set
            Me._Ix = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Iy")>
    Public Property Iy() As Double?
        Get
            Return Me._Iy
        End Get
        Set
            Me._Iy = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Lu")>
    Public Property Lu() As Double?
        Get
            Return Me._Lu
        End Get
        Set
            Me._Lu = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Kx")>
    Public Property Kx() As Double?
        Get
            Return Me._Kx
        End Get
        Set
            Me._Kx = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Ky")>
    Public Property Ky() As Double?
        Get
            Return Me._Ky
        End Get
        Set
            Me._Ky = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt Hole Size")>
    Public Property bolt_hole_size() As Double?
        Get
            Return Me._bolt_hole_size
        End Get
        Set
            Me._bolt_hole_size = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Area Net")>
    Public Property area_net() As Double?
        Get
            Return Me._area_net
        End Get
        Set
            Me._area_net = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Shear Lag")>
    Public Property shear_lag() As Double?
        Get
            Return Me._shear_lag
        End Get
        Set
            Me._shear_lag = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Connection Type Bot")>
    Public Property connection_type_bot() As String
        Get
            Return Me._connection_type_bot
        End Get
        Set
            Me._connection_type_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Connection Cap Revf Bot")>
    Public Property connection_cap_revF_bot() As Double?
        Get
            Return Me._connection_cap_revF_bot
        End Get
        Set
            Me._connection_cap_revF_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Connection Cap Revg Bot")>
    Public Property connection_cap_revG_bot() As Double?
        Get
            Return Me._connection_cap_revG_bot
        End Get
        Set
            Me._connection_cap_revG_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Connection Cap Revh Bot")>
    Public Property connection_cap_revH_bot() As Double?
        Get
            Return Me._connection_cap_revH_bot
        End Get
        Set
            Me._connection_cap_revH_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt Id Bot")>
    Public Property bolt_id_bot() As Integer?
        Get
            Return Me._bolt_id_bot
        End Get
        Set
            Me._bolt_id_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Local Bolt Id Bot")>
    Public Property local_bolt_id_bot() As Integer?
        Get
            Return Me._local_bolt_id_bot
        End Get
        Set
            Me._local_bolt_id_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt N Or X Bot")>
    Public Property bolt_N_or_X_bot() As String
        Get
            Return Me._bolt_N_or_X_bot
        End Get
        Set
            Me._bolt_N_or_X_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt Num Bot")>
    Public Property bolt_num_bot() As Integer?
        Get
            Return Me._bolt_num_bot
        End Get
        Set
            Me._bolt_num_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt Spacing Bot")>
    Public Property bolt_spacing_bot() As Double?
        Get
            Return Me._bolt_spacing_bot
        End Get
        Set
            Me._bolt_spacing_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt Edge Dist Bot")>
    Public Property bolt_edge_dist_bot() As Double?
        Get
            Return Me._bolt_edge_dist_bot
        End Get
        Set
            Me._bolt_edge_dist_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Flangeorbp Connected Bot")>
    Public Property FlangeOrBP_connected_bot() As Boolean?
        Get
            Return Me._FlangeOrBP_connected_bot
        End Get
        Set
            Me._FlangeOrBP_connected_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Grade Bot")>
    Public Property weld_grade_bot() As Double?
        Get
            Return Me._weld_grade_bot
        End Get
        Set
            Me._weld_grade_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Trans Type Bot")>
    Public Property weld_trans_type_bot() As String
        Get
            Return Me._weld_trans_type_bot
        End Get
        Set
            Me._weld_trans_type_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Trans Length Bot")>
    Public Property weld_trans_length_bot() As Double?
        Get
            Return Me._weld_trans_length_bot
        End Get
        Set
            Me._weld_trans_length_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Groove Depth Bot")>
    Public Property weld_groove_depth_bot() As Double?
        Get
            Return Me._weld_groove_depth_bot
        End Get
        Set
            Me._weld_groove_depth_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Groove Angle Bot")>
    Public Property weld_groove_angle_bot() As Integer?
        Get
            Return Me._weld_groove_angle_bot
        End Get
        Set
            Me._weld_groove_angle_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Trans Fillet Size Bot")>
    Public Property weld_trans_fillet_size_bot() As Double?
        Get
            Return Me._weld_trans_fillet_size_bot
        End Get
        Set
            Me._weld_trans_fillet_size_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Trans Eff Throat Bot")>
    Public Property weld_trans_eff_throat_bot() As Double?
        Get
            Return Me._weld_trans_eff_throat_bot
        End Get
        Set
            Me._weld_trans_eff_throat_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Long Type Bot")>
    Public Property weld_long_type_bot() As String
        Get
            Return Me._weld_long_type_bot
        End Get
        Set
            Me._weld_long_type_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Long Length Bot")>
    Public Property weld_long_length_bot() As Double?
        Get
            Return Me._weld_long_length_bot
        End Get
        Set
            Me._weld_long_length_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Long Fillet Size Bot")>
    Public Property weld_long_fillet_size_bot() As Double?
        Get
            Return Me._weld_long_fillet_size_bot
        End Get
        Set
            Me._weld_long_fillet_size_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Long Eff Throat Bot")>
    Public Property weld_long_eff_throat_bot() As Double?
        Get
            Return Me._weld_long_eff_throat_bot
        End Get
        Set
            Me._weld_long_eff_throat_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Top Bot Connections Symmetrical")>
    Public Property top_bot_connections_symmetrical() As Boolean?
        Get
            Return Me._top_bot_connections_symmetrical
        End Get
        Set
            Me._top_bot_connections_symmetrical = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Connection Type Top")>
    Public Property connection_type_top() As String
        Get
            Return Me._connection_type_top
        End Get
        Set
            Me._connection_type_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Connection Cap Revf Top")>
    Public Property connection_cap_revF_top() As Double?
        Get
            Return Me._connection_cap_revF_top
        End Get
        Set
            Me._connection_cap_revF_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Connection Cap Revg Top")>
    Public Property connection_cap_revG_top() As Double?
        Get
            Return Me._connection_cap_revG_top
        End Get
        Set
            Me._connection_cap_revG_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Connection Cap Revh Top")>
    Public Property connection_cap_revH_top() As Double?
        Get
            Return Me._connection_cap_revH_top
        End Get
        Set
            Me._connection_cap_revH_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt Id Top")>
    Public Property bolt_id_top() As Integer?
        Get
            Return Me._bolt_id_top
        End Get
        Set
            Me._bolt_id_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Local Bolt Id Top")>
    Public Property local_bolt_id_top() As Integer?
        Get
            Return Me._local_bolt_id_top
        End Get
        Set
            Me._local_bolt_id_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt N Or X Top")>
    Public Property bolt_N_or_X_top() As String
        Get
            Return Me._bolt_N_or_X_top
        End Get
        Set
            Me._bolt_N_or_X_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt Num Top")>
    Public Property bolt_num_top() As Integer?
        Get
            Return Me._bolt_num_top
        End Get
        Set
            Me._bolt_num_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt Spacing Top")>
    Public Property bolt_spacing_top() As Double?
        Get
            Return Me._bolt_spacing_top
        End Get
        Set
            Me._bolt_spacing_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Bolt Edge Dist Top")>
    Public Property bolt_edge_dist_top() As Double?
        Get
            Return Me._bolt_edge_dist_top
        End Get
        Set
            Me._bolt_edge_dist_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Flangeorbp Connected Top")>
    Public Property FlangeOrBP_connected_top() As Boolean?
        Get
            Return Me._FlangeOrBP_connected_top
        End Get
        Set
            Me._FlangeOrBP_connected_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Grade Top")>
    Public Property weld_grade_top() As Double?
        Get
            Return Me._weld_grade_top
        End Get
        Set
            Me._weld_grade_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Trans Type Top")>
    Public Property weld_trans_type_top() As String
        Get
            Return Me._weld_trans_type_top
        End Get
        Set
            Me._weld_trans_type_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Trans Length Top")>
    Public Property weld_trans_length_top() As Double?
        Get
            Return Me._weld_trans_length_top
        End Get
        Set
            Me._weld_trans_length_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Groove Depth Top")>
    Public Property weld_groove_depth_top() As Double?
        Get
            Return Me._weld_groove_depth_top
        End Get
        Set
            Me._weld_groove_depth_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Groove Angle Top")>
    Public Property weld_groove_angle_top() As Integer?
        Get
            Return Me._weld_groove_angle_top
        End Get
        Set
            Me._weld_groove_angle_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Trans Fillet Size Top")>
    Public Property weld_trans_fillet_size_top() As Double?
        Get
            Return Me._weld_trans_fillet_size_top
        End Get
        Set
            Me._weld_trans_fillet_size_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Trans Eff Throat Top")>
    Public Property weld_trans_eff_throat_top() As Double?
        Get
            Return Me._weld_trans_eff_throat_top
        End Get
        Set
            Me._weld_trans_eff_throat_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Long Type Top")>
    Public Property weld_long_type_top() As String
        Get
            Return Me._weld_long_type_top
        End Get
        Set
            Me._weld_long_type_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Long Length Top")>
    Public Property weld_long_length_top() As Double?
        Get
            Return Me._weld_long_length_top
        End Get
        Set
            Me._weld_long_length_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Long Fillet Size Top")>
    Public Property weld_long_fillet_size_top() As Double?
        Get
            Return Me._weld_long_fillet_size_top
        End Get
        Set
            Me._weld_long_fillet_size_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Weld Long Eff Throat Top")>
    Public Property weld_long_eff_throat_top() As Double?
        Get
            Return Me._weld_long_eff_throat_top
        End Get
        Set
            Me._weld_long_eff_throat_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Conn Length Channel")>
    Public Property conn_length_channel() As Double?
        Get
            Return Me._conn_length_channel
        End Get
        Set
            Me._conn_length_channel = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Conn Length Bot")>
    Public Property conn_length_bot() As Double?
        Get
            Return Me._conn_length_bot
        End Get
        Set
            Me._conn_length_bot = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Conn Length Top")>
    Public Property conn_length_top() As Double?
        Get
            Return Me._conn_length_top
        End Get
        Set
            Me._conn_length_top = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Comp Xx F")>
    Public Property cap_comp_xx_f() As Double?
        Get
            Return Me._cap_comp_xx_f
        End Get
        Set
            Me._cap_comp_xx_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Comp Yy F")>
    Public Property cap_comp_yy_f() As Double?
        Get
            Return Me._cap_comp_yy_f
        End Get
        Set
            Me._cap_comp_yy_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Tens Yield F")>
    Public Property cap_tens_yield_f() As Double?
        Get
            Return Me._cap_tens_yield_f
        End Get
        Set
            Me._cap_tens_yield_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Tens Rupture F")>
    Public Property cap_tens_rupture_f() As Double?
        Get
            Return Me._cap_tens_rupture_f
        End Get
        Set
            Me._cap_tens_rupture_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Shear F")>
    Public Property cap_shear_f() As Double?
        Get
            Return Me._cap_shear_f
        End Get
        Set
            Me._cap_shear_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Bolt Shear Bot F")>
    Public Property cap_bolt_shear_bot_f() As Double?
        Get
            Return Me._cap_bolt_shear_bot_f
        End Get
        Set
            Me._cap_bolt_shear_bot_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Bolt Shear Top F")>
    Public Property cap_bolt_shear_top_f() As Double?
        Get
            Return Me._cap_bolt_shear_top_f
        End Get
        Set
            Me._cap_bolt_shear_top_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Nodeform Bot F")>
    Public Property cap_boltshaft_bearing_nodeform_bot_f() As Double?
        Get
            Return Me._cap_boltshaft_bearing_nodeform_bot_f
        End Get
        Set
            Me._cap_boltshaft_bearing_nodeform_bot_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Deform Bot F")>
    Public Property cap_boltshaft_bearing_deform_bot_f() As Double?
        Get
            Return Me._cap_boltshaft_bearing_deform_bot_f
        End Get
        Set
            Me._cap_boltshaft_bearing_deform_bot_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Nodeform Top F")>
    Public Property cap_boltshaft_bearing_nodeform_top_f() As Double?
        Get
            Return Me._cap_boltshaft_bearing_nodeform_top_f
        End Get
        Set
            Me._cap_boltshaft_bearing_nodeform_top_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Deform Top F")>
    Public Property cap_boltshaft_bearing_deform_top_f() As Double?
        Get
            Return Me._cap_boltshaft_bearing_deform_top_f
        End Get
        Set
            Me._cap_boltshaft_bearing_deform_top_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Nodeform Bot F")>
    Public Property cap_boltreinf_bearing_nodeform_bot_f() As Double?
        Get
            Return Me._cap_boltreinf_bearing_nodeform_bot_f
        End Get
        Set
            Me._cap_boltreinf_bearing_nodeform_bot_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Deform Bot F")>
    Public Property cap_boltreinf_bearing_deform_bot_f() As Double?
        Get
            Return Me._cap_boltreinf_bearing_deform_bot_f
        End Get
        Set
            Me._cap_boltreinf_bearing_deform_bot_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Nodeform Top F")>
    Public Property cap_boltreinf_bearing_nodeform_top_f() As Double?
        Get
            Return Me._cap_boltreinf_bearing_nodeform_top_f
        End Get
        Set
            Me._cap_boltreinf_bearing_nodeform_top_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Deform Top F")>
    Public Property cap_boltreinf_bearing_deform_top_f() As Double?
        Get
            Return Me._cap_boltreinf_bearing_deform_top_f
        End Get
        Set
            Me._cap_boltreinf_bearing_deform_top_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Trans Bot F")>
    Public Property cap_weld_trans_bot_f() As Double?
        Get
            Return Me._cap_weld_trans_bot_f
        End Get
        Set
            Me._cap_weld_trans_bot_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Long Bot F")>
    Public Property cap_weld_long_bot_f() As Double?
        Get
            Return Me._cap_weld_long_bot_f
        End Get
        Set
            Me._cap_weld_long_bot_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Trans Top F")>
    Public Property cap_weld_trans_top_f() As Double?
        Get
            Return Me._cap_weld_trans_top_f
        End Get
        Set
            Me._cap_weld_trans_top_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Long Top F")>
    Public Property cap_weld_long_top_f() As Double?
        Get
            Return Me._cap_weld_long_top_f
        End Get
        Set
            Me._cap_weld_long_top_f = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Comp Xx G")>
    Public Property cap_comp_xx_g() As Double?
        Get
            Return Me._cap_comp_xx_g
        End Get
        Set
            Me._cap_comp_xx_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Comp Yy G")>
    Public Property cap_comp_yy_g() As Double?
        Get
            Return Me._cap_comp_yy_g
        End Get
        Set
            Me._cap_comp_yy_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Tens Yield G")>
    Public Property cap_tens_yield_g() As Double?
        Get
            Return Me._cap_tens_yield_g
        End Get
        Set
            Me._cap_tens_yield_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Tens Rupture G")>
    Public Property cap_tens_rupture_g() As Double?
        Get
            Return Me._cap_tens_rupture_g
        End Get
        Set
            Me._cap_tens_rupture_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Shear G")>
    Public Property cap_shear_g() As Double?
        Get
            Return Me._cap_shear_g
        End Get
        Set
            Me._cap_shear_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Bolt Shear Bot G")>
    Public Property cap_bolt_shear_bot_g() As Double?
        Get
            Return Me._cap_bolt_shear_bot_g
        End Get
        Set
            Me._cap_bolt_shear_bot_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Bolt Shear Top G")>
    Public Property cap_bolt_shear_top_g() As Double?
        Get
            Return Me._cap_bolt_shear_top_g
        End Get
        Set
            Me._cap_bolt_shear_top_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Nodeform Bot G")>
    Public Property cap_boltshaft_bearing_nodeform_bot_g() As Double?
        Get
            Return Me._cap_boltshaft_bearing_nodeform_bot_g
        End Get
        Set
            Me._cap_boltshaft_bearing_nodeform_bot_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Deform Bot G")>
    Public Property cap_boltshaft_bearing_deform_bot_g() As Double?
        Get
            Return Me._cap_boltshaft_bearing_deform_bot_g
        End Get
        Set
            Me._cap_boltshaft_bearing_deform_bot_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Nodeform Top G")>
    Public Property cap_boltshaft_bearing_nodeform_top_g() As Double?
        Get
            Return Me._cap_boltshaft_bearing_nodeform_top_g
        End Get
        Set
            Me._cap_boltshaft_bearing_nodeform_top_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Deform Top G")>
    Public Property cap_boltshaft_bearing_deform_top_g() As Double?
        Get
            Return Me._cap_boltshaft_bearing_deform_top_g
        End Get
        Set
            Me._cap_boltshaft_bearing_deform_top_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Nodeform Bot G")>
    Public Property cap_boltreinf_bearing_nodeform_bot_g() As Double?
        Get
            Return Me._cap_boltreinf_bearing_nodeform_bot_g
        End Get
        Set
            Me._cap_boltreinf_bearing_nodeform_bot_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Deform Bot G")>
    Public Property cap_boltreinf_bearing_deform_bot_g() As Double?
        Get
            Return Me._cap_boltreinf_bearing_deform_bot_g
        End Get
        Set
            Me._cap_boltreinf_bearing_deform_bot_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Nodeform Top G")>
    Public Property cap_boltreinf_bearing_nodeform_top_g() As Double?
        Get
            Return Me._cap_boltreinf_bearing_nodeform_top_g
        End Get
        Set
            Me._cap_boltreinf_bearing_nodeform_top_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Deform Top G")>
    Public Property cap_boltreinf_bearing_deform_top_g() As Double?
        Get
            Return Me._cap_boltreinf_bearing_deform_top_g
        End Get
        Set
            Me._cap_boltreinf_bearing_deform_top_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Trans Bot G")>
    Public Property cap_weld_trans_bot_g() As Double?
        Get
            Return Me._cap_weld_trans_bot_g
        End Get
        Set
            Me._cap_weld_trans_bot_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Long Bot G")>
    Public Property cap_weld_long_bot_g() As Double?
        Get
            Return Me._cap_weld_long_bot_g
        End Get
        Set
            Me._cap_weld_long_bot_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Trans Top G")>
    Public Property cap_weld_trans_top_g() As Double?
        Get
            Return Me._cap_weld_trans_top_g
        End Get
        Set
            Me._cap_weld_trans_top_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Long Top G")>
    Public Property cap_weld_long_top_g() As Double?
        Get
            Return Me._cap_weld_long_top_g
        End Get
        Set
            Me._cap_weld_long_top_g = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Comp Xx H")>
    Public Property cap_comp_xx_h() As Double?
        Get
            Return Me._cap_comp_xx_h
        End Get
        Set
            Me._cap_comp_xx_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Comp Yy H")>
    Public Property cap_comp_yy_h() As Double?
        Get
            Return Me._cap_comp_yy_h
        End Get
        Set
            Me._cap_comp_yy_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Tens Yield H")>
    Public Property cap_tens_yield_h() As Double?
        Get
            Return Me._cap_tens_yield_h
        End Get
        Set
            Me._cap_tens_yield_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Tens Rupture H")>
    Public Property cap_tens_rupture_h() As Double?
        Get
            Return Me._cap_tens_rupture_h
        End Get
        Set
            Me._cap_tens_rupture_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Shear H")>
    Public Property cap_shear_h() As Double?
        Get
            Return Me._cap_shear_h
        End Get
        Set
            Me._cap_shear_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Bolt Shear Bot H")>
    Public Property cap_bolt_shear_bot_h() As Double?
        Get
            Return Me._cap_bolt_shear_bot_h
        End Get
        Set
            Me._cap_bolt_shear_bot_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Bolt Shear Top H")>
    Public Property cap_bolt_shear_top_h() As Double?
        Get
            Return Me._cap_bolt_shear_top_h
        End Get
        Set
            Me._cap_bolt_shear_top_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Nodeform Bot H")>
    Public Property cap_boltshaft_bearing_nodeform_bot_h() As Double?
        Get
            Return Me._cap_boltshaft_bearing_nodeform_bot_h
        End Get
        Set
            Me._cap_boltshaft_bearing_nodeform_bot_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Deform Bot H")>
    Public Property cap_boltshaft_bearing_deform_bot_h() As Double?
        Get
            Return Me._cap_boltshaft_bearing_deform_bot_h
        End Get
        Set
            Me._cap_boltshaft_bearing_deform_bot_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Nodeform Top H")>
    Public Property cap_boltshaft_bearing_nodeform_top_h() As Double?
        Get
            Return Me._cap_boltshaft_bearing_nodeform_top_h
        End Get
        Set
            Me._cap_boltshaft_bearing_nodeform_top_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltshaft Bearing Deform Top H")>
    Public Property cap_boltshaft_bearing_deform_top_h() As Double?
        Get
            Return Me._cap_boltshaft_bearing_deform_top_h
        End Get
        Set
            Me._cap_boltshaft_bearing_deform_top_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Nodeform Bot H")>
    Public Property cap_boltreinf_bearing_nodeform_bot_h() As Double?
        Get
            Return Me._cap_boltreinf_bearing_nodeform_bot_h
        End Get
        Set
            Me._cap_boltreinf_bearing_nodeform_bot_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Deform Bot H")>
    Public Property cap_boltreinf_bearing_deform_bot_h() As Double?
        Get
            Return Me._cap_boltreinf_bearing_deform_bot_h
        End Get
        Set
            Me._cap_boltreinf_bearing_deform_bot_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Nodeform Top H")>
    Public Property cap_boltreinf_bearing_nodeform_top_h() As Double?
        Get
            Return Me._cap_boltreinf_bearing_nodeform_top_h
        End Get
        Set
            Me._cap_boltreinf_bearing_nodeform_top_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Boltreinf Bearing Deform Top H")>
    Public Property cap_boltreinf_bearing_deform_top_h() As Double?
        Get
            Return Me._cap_boltreinf_bearing_deform_top_h
        End Get
        Set
            Me._cap_boltreinf_bearing_deform_top_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Trans Bot H")>
    Public Property cap_weld_trans_bot_h() As Double?
        Get
            Return Me._cap_weld_trans_bot_h
        End Get
        Set
            Me._cap_weld_trans_bot_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Long Bot H")>
    Public Property cap_weld_long_bot_h() As Double?
        Get
            Return Me._cap_weld_long_bot_h
        End Get
        Set
            Me._cap_weld_long_bot_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Trans Top H")>
    Public Property cap_weld_trans_top_h() As Double?
        Get
            Return Me._cap_weld_trans_top_h
        End Get
        Set
            Me._cap_weld_trans_top_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Cap Weld Long Top H")>
    Public Property cap_weld_long_top_h() As Double?
        Get
            Return Me._cap_weld_long_top_h
        End Get
        Set
            Me._cap_weld_long_top_h = Value
        End Set
    End Property
    <Category("CCIpole Reinf DB"), Description(""), DisplayName("Ind Default")>
    Public Property ind_default() As Boolean?
        Get
            Return Me._ind_default
        End Get
        Set
            Me._ind_default = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal Row As DataRow, Optional ByVal Parent As EDSObject = Nothing) ', ByRef strDS As DataSet
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet
        Dim dr = Row

        Me.reinf_id = DBtoNullableInt(dr.Item("ID"))
        Me.local_reinf_id = DBtoNullableInt(dr.Item("local_reinf_id"))
        Me.name = DBtoStr(dr.Item("name"))
        Me.type = DBtoStr(dr.Item("type"))
        Me.b = DBtoNullableDbl(dr.Item("b"))
        Me.h = DBtoNullableDbl(dr.Item("h"))
        Me.sr_diam = DBtoNullableDbl(dr.Item("sr_diam"))
        Me.channel_thkns_web = DBtoNullableDbl(dr.Item("channel_thkns_web"))
        Me.channel_thkns_flange = DBtoNullableDbl(dr.Item("channel_thkns_flange"))
        Me.channel_eo = DBtoNullableDbl(dr.Item("channel_eo"))
        Me.channel_J = DBtoNullableDbl(dr.Item("channel_J"))
        Me.channel_Cw = DBtoNullableDbl(dr.Item("channel_Cw"))
        Me.area_gross = DBtoNullableDbl(dr.Item("area_gross"))
        Me.centroid = DBtoNullableDbl(dr.Item("centroid"))
        Me.istension = DBtoNullableBool(dr.Item("istension"))
        Me.matl_id = DBtoNullableInt(dr.Item("matl_id"))
        Me.local_matl_id = DBtoNullableInt(dr.Item("local_matl_id"))
        Me.Ix = DBtoNullableDbl(dr.Item("Ix"))
        Me.Iy = DBtoNullableDbl(dr.Item("Iy"))
        Me.Lu = DBtoNullableDbl(dr.Item("Lu"))
        Me.Kx = DBtoNullableDbl(dr.Item("Kx"))
        Me.Ky = DBtoNullableDbl(dr.Item("Ky"))
        Me.bolt_hole_size = DBtoNullableDbl(dr.Item("bolt_hole_size"))
        Me.area_net = DBtoNullableDbl(dr.Item("area_net"))
        Me.shear_lag = DBtoNullableDbl(dr.Item("shear_lag"))
        Me.connection_type_bot = DBtoStr(dr.Item("connection_type_bot"))
        Me.connection_cap_revF_bot = DBtoNullableDbl(dr.Item("connection_cap_revF_bot"))
        Me.connection_cap_revG_bot = DBtoNullableDbl(dr.Item("connection_cap_revG_bot"))
        Me.connection_cap_revH_bot = DBtoNullableDbl(dr.Item("connection_cap_revH_bot"))
        Me.bolt_id_bot = DBtoNullableInt(dr.Item("bolt_id_bot"))
        Me.local_bolt_id_bot = DBtoNullableInt(dr.Item("local_bolt_id_bot"))
        If DBtoStr(dr.Item("bolt_N_or_X_bot")) = "-" Or DBtoStr(dr.Item("bolt_N_or_X_bot")) = "n/a" Then
            Me.bolt_N_or_X_bot = ""
        Else
            Me.bolt_N_or_X_bot = DBtoStr(dr.Item("bolt_N_or_X_bot"))
        End If
        Me.bolt_num_bot = DBtoNullableInt(dr.Item("bolt_num_bot"))
        Me.bolt_spacing_bot = DBtoNullableDbl(dr.Item("bolt_spacing_bot"))
        Me.bolt_edge_dist_bot = DBtoNullableDbl(dr.Item("bolt_edge_dist_bot"))
        Me.FlangeOrBP_connected_bot = DBtoNullableBool(dr.Item("FlangeOrBP_connected_bot"))
        Me.weld_grade_bot = DBtoNullableDbl(dr.Item("weld_grade_bot"))
        If DBtoStr(dr.Item("weld_trans_type_bot")) = "-" Or DBtoStr(dr.Item("weld_trans_type_bot")) = "n/a" Then
            Me.weld_trans_type_bot = ""
        Else
            Me.weld_trans_type_bot = DBtoStr(dr.Item("weld_trans_type_bot"))
        End If
        Me.weld_trans_length_bot = DBtoNullableDbl(dr.Item("weld_trans_length_bot"))
        Me.weld_groove_depth_bot = DBtoNullableDbl(dr.Item("weld_groove_depth_bot"))
        Me.weld_groove_angle_bot = DBtoNullableInt(dr.Item("weld_groove_angle_bot"))
        Me.weld_trans_fillet_size_bot = DBtoNullableDbl(dr.Item("weld_trans_fillet_size_bot"))
        Me.weld_trans_eff_throat_bot = DBtoNullableDbl(dr.Item("weld_trans_eff_throat_bot"))
        If DBtoStr(dr.Item("weld_long_type_bot")) = "-" Or DBtoStr(dr.Item("weld_long_type_bot")) = "n/a" Then
            Me.weld_long_type_bot = ""
        Else
            Me.weld_long_type_bot = DBtoStr(dr.Item("weld_long_type_bot"))
        End If
        Me.weld_long_length_bot = DBtoNullableDbl(dr.Item("weld_long_length_bot"))
        Me.weld_long_fillet_size_bot = DBtoNullableDbl(dr.Item("weld_long_fillet_size_bot"))
        Me.weld_long_eff_throat_bot = DBtoNullableDbl(dr.Item("weld_long_eff_throat_bot"))
        Me.top_bot_connections_symmetrical = DBtoNullableBool(dr.Item("top_bot_connections_symmetrical"))
        Me.connection_type_top = DBtoStr(dr.Item("connection_type_top"))
        Me.connection_cap_revF_top = DBtoNullableDbl(dr.Item("connection_cap_revF_top"))
        Me.connection_cap_revG_top = DBtoNullableDbl(dr.Item("connection_cap_revG_top"))
        Me.connection_cap_revH_top = DBtoNullableDbl(dr.Item("connection_cap_revH_top"))
        Me.bolt_id_top = DBtoNullableInt(dr.Item("bolt_id_top"))
        Me.local_bolt_id_top = DBtoNullableInt(dr.Item("local_bolt_id_top"))
        If DBtoStr(dr.Item("bolt_N_or_X_top")) = "-" Or DBtoStr(dr.Item("bolt_N_or_X_top")) = "n/a" Then
            Me.bolt_N_or_X_top = ""
        Else
            Me.bolt_N_or_X_top = DBtoStr(dr.Item("bolt_N_or_X_top"))
        End If
        Me.bolt_num_top = DBtoNullableInt(dr.Item("bolt_num_top"))
        Me.bolt_spacing_top = DBtoNullableDbl(dr.Item("bolt_spacing_top"))
        Me.bolt_edge_dist_top = DBtoNullableDbl(dr.Item("bolt_edge_dist_top"))
        Me.FlangeOrBP_connected_top = DBtoNullableBool(dr.Item("FlangeOrBP_connected_top"))
        Me.weld_grade_top = DBtoNullableDbl(dr.Item("weld_grade_top"))
        If DBtoStr(dr.Item("weld_trans_type_top")) = "-" Or DBtoStr(dr.Item("weld_trans_type_top")) = "n/a" Then
            Me.weld_trans_type_top = ""
        Else
            Me.weld_trans_type_top = DBtoStr(dr.Item("weld_trans_type_top"))
        End If
        Me.weld_trans_length_top = DBtoNullableDbl(dr.Item("weld_trans_length_top"))
        Me.weld_groove_depth_top = DBtoNullableDbl(dr.Item("weld_groove_depth_top"))
        Me.weld_groove_angle_top = DBtoNullableInt(dr.Item("weld_groove_angle_top"))
        Me.weld_trans_fillet_size_top = DBtoNullableDbl(dr.Item("weld_trans_fillet_size_top"))
        Me.weld_trans_eff_throat_top = DBtoNullableDbl(dr.Item("weld_trans_eff_throat_top"))
        If DBtoStr(dr.Item("weld_long_type_top")) = "-" Or DBtoStr(dr.Item("weld_long_type_top")) = "n/a" Then
            Me.weld_long_type_top = ""
        Else
            Me.weld_long_type_top = DBtoStr(dr.Item("weld_long_type_top"))
        End If
        Me.weld_long_length_top = DBtoNullableDbl(dr.Item("weld_long_length_top"))
        Me.weld_long_fillet_size_top = DBtoNullableDbl(dr.Item("weld_long_fillet_size_top"))
        Me.weld_long_eff_throat_top = DBtoNullableDbl(dr.Item("weld_long_eff_throat_top"))
        Me.conn_length_channel = DBtoNullableDbl(dr.Item("conn_length_channel"))
        Me.conn_length_bot = DBtoNullableDbl(dr.Item("conn_length_bot"))
        Me.conn_length_top = DBtoNullableDbl(dr.Item("conn_length_top"))
        Me.cap_comp_xx_f = DBtoNullableDbl(dr.Item("cap_comp_xx_f"))
        Me.cap_comp_yy_f = DBtoNullableDbl(dr.Item("cap_comp_yy_f"))
        Me.cap_tens_yield_f = DBtoNullableDbl(dr.Item("cap_tens_yield_f"))
        Me.cap_tens_rupture_f = DBtoNullableDbl(dr.Item("cap_tens_rupture_f"))
        Me.cap_shear_f = DBtoNullableDbl(dr.Item("cap_shear_f"))
        Me.cap_bolt_shear_bot_f = DBtoNullableDbl(dr.Item("cap_bolt_shear_bot_f"))
        Me.cap_bolt_shear_top_f = DBtoNullableDbl(dr.Item("cap_bolt_shear_top_f"))
        Me.cap_boltshaft_bearing_nodeform_bot_f = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_nodeform_bot_f"))
        Me.cap_boltshaft_bearing_deform_bot_f = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_deform_bot_f"))
        Me.cap_boltshaft_bearing_nodeform_top_f = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_nodeform_top_f"))
        Me.cap_boltshaft_bearing_deform_top_f = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_deform_top_f"))
        Me.cap_boltreinf_bearing_nodeform_bot_f = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_nodeform_bot_f"))
        Me.cap_boltreinf_bearing_deform_bot_f = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_deform_bot_f"))
        Me.cap_boltreinf_bearing_nodeform_top_f = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_nodeform_top_f"))
        Me.cap_boltreinf_bearing_deform_top_f = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_deform_top_f"))
        Me.cap_weld_trans_bot_f = DBtoNullableDbl(dr.Item("cap_weld_trans_bot_f"))
        Me.cap_weld_long_bot_f = DBtoNullableDbl(dr.Item("cap_weld_long_bot_f"))
        Me.cap_weld_trans_top_f = DBtoNullableDbl(dr.Item("cap_weld_trans_top_f"))
        Me.cap_weld_long_top_f = DBtoNullableDbl(dr.Item("cap_weld_long_top_f"))
        Me.cap_comp_xx_g = DBtoNullableDbl(dr.Item("cap_comp_xx_g"))
        Me.cap_comp_yy_g = DBtoNullableDbl(dr.Item("cap_comp_yy_g"))
        Me.cap_tens_yield_g = DBtoNullableDbl(dr.Item("cap_tens_yield_g"))
        Me.cap_tens_rupture_g = DBtoNullableDbl(dr.Item("cap_tens_rupture_g"))
        Me.cap_shear_g = DBtoNullableDbl(dr.Item("cap_shear_g"))
        Me.cap_bolt_shear_bot_g = DBtoNullableDbl(dr.Item("cap_bolt_shear_bot_g"))
        Me.cap_bolt_shear_top_g = DBtoNullableDbl(dr.Item("cap_bolt_shear_top_g"))
        Me.cap_boltshaft_bearing_nodeform_bot_g = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_nodeform_bot_g"))
        Me.cap_boltshaft_bearing_deform_bot_g = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_deform_bot_g"))
        Me.cap_boltshaft_bearing_nodeform_top_g = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_nodeform_top_g"))
        Me.cap_boltshaft_bearing_deform_top_g = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_deform_top_g"))
        Me.cap_boltreinf_bearing_nodeform_bot_g = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_nodeform_bot_g"))
        Me.cap_boltreinf_bearing_deform_bot_g = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_deform_bot_g"))
        Me.cap_boltreinf_bearing_nodeform_top_g = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_nodeform_top_g"))
        Me.cap_boltreinf_bearing_deform_top_g = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_deform_top_g"))
        Me.cap_weld_trans_bot_g = DBtoNullableDbl(dr.Item("cap_weld_trans_bot_g"))
        Me.cap_weld_long_bot_g = DBtoNullableDbl(dr.Item("cap_weld_long_bot_g"))
        Me.cap_weld_trans_top_g = DBtoNullableDbl(dr.Item("cap_weld_trans_top_g"))
        Me.cap_weld_long_top_g = DBtoNullableDbl(dr.Item("cap_weld_long_top_g"))
        Me.cap_comp_xx_h = DBtoNullableDbl(dr.Item("cap_comp_xx_h"))
        Me.cap_comp_yy_h = DBtoNullableDbl(dr.Item("cap_comp_yy_h"))
        Me.cap_tens_yield_h = DBtoNullableDbl(dr.Item("cap_tens_yield_h"))
        Me.cap_tens_rupture_h = DBtoNullableDbl(dr.Item("cap_tens_rupture_h"))
        Me.cap_shear_h = DBtoNullableDbl(dr.Item("cap_shear_h"))
        Me.cap_bolt_shear_bot_h = DBtoNullableDbl(dr.Item("cap_bolt_shear_bot_h"))
        Me.cap_bolt_shear_top_h = DBtoNullableDbl(dr.Item("cap_bolt_shear_top_h"))
        Me.cap_boltshaft_bearing_nodeform_bot_h = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_nodeform_bot_h"))
        Me.cap_boltshaft_bearing_deform_bot_h = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_deform_bot_h"))
        Me.cap_boltshaft_bearing_nodeform_top_h = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_nodeform_top_h"))
        Me.cap_boltshaft_bearing_deform_top_h = DBtoNullableDbl(dr.Item("cap_boltshaft_bearing_deform_top_h"))
        Me.cap_boltreinf_bearing_nodeform_bot_h = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_nodeform_bot_h"))
        Me.cap_boltreinf_bearing_deform_bot_h = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_deform_bot_h"))
        Me.cap_boltreinf_bearing_nodeform_top_h = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_nodeform_top_h"))
        Me.cap_boltreinf_bearing_deform_top_h = DBtoNullableDbl(dr.Item("cap_boltreinf_bearing_deform_top_h"))
        Me.cap_weld_trans_bot_h = DBtoNullableDbl(dr.Item("cap_weld_trans_bot_h"))
        Me.cap_weld_long_bot_h = DBtoNullableDbl(dr.Item("cap_weld_long_bot_h"))
        Me.cap_weld_trans_top_h = DBtoNullableDbl(dr.Item("cap_weld_trans_top_h"))
        Me.cap_weld_long_top_h = DBtoNullableDbl(dr.Item("cap_weld_long_top_h"))
        Me.ind_default = DBtoNullableBool(dr.Item("ind_default"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.reinf_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID") '(Me.pole_id.ToString.FormatDBValue) - Does pole_id need to be added? - MRR
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_reinf_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.name.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.b.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.sr_diam.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.channel_thkns_web.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.channel_thkns_flange.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.channel_eo.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.channel_J.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.channel_Cw.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.area_gross.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.centroid.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.istension.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel4ID") '(Me.matl_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_matl_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Ix.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Iy.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Lu.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Kx.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Ky.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_hole_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.area_net.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.shear_lag.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_type_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_cap_revF_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_cap_revG_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_cap_revH_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@BotBoltID") '("@SubLevel3ID") '(Me.bolt_id_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_bolt_id_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_N_or_X_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_num_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_spacing_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_edge_dist_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FlangeOrBP_connected_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_grade_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_trans_type_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_trans_length_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_groove_depth_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_groove_angle_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_trans_fillet_size_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_trans_eff_throat_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_long_type_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_long_length_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_long_fillet_size_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_long_eff_throat_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.top_bot_connections_symmetrical.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_type_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_cap_revF_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_cap_revG_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_cap_revH_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopBoltID") '("@SubLevel3ID") '(Me.bolt_id_top.ToString.FormatDBValue) - I think this needs to be a new variable - MRR
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_bolt_id_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_N_or_X_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_num_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_spacing_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_edge_dist_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FlangeOrBP_connected_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_grade_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_trans_type_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_trans_length_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_groove_depth_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_groove_angle_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_trans_fillet_size_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_trans_eff_throat_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_long_type_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_long_length_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_long_fillet_size_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_long_eff_throat_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.conn_length_channel.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.conn_length_bot.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.conn_length_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_comp_xx_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_comp_yy_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_tens_yield_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_tens_rupture_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_shear_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_bolt_shear_bot_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_bolt_shear_top_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_nodeform_bot_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_deform_bot_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_nodeform_top_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_deform_top_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_nodeform_bot_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_deform_bot_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_nodeform_top_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_deform_top_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_trans_bot_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_long_bot_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_trans_top_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_long_top_f.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_comp_xx_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_comp_yy_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_tens_yield_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_tens_rupture_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_shear_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_bolt_shear_bot_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_bolt_shear_top_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_nodeform_bot_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_deform_bot_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_nodeform_top_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_deform_top_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_nodeform_bot_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_deform_bot_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_nodeform_top_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_deform_top_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_trans_bot_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_long_bot_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_trans_top_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_long_top_g.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_comp_xx_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_comp_yy_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_tens_yield_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_tens_rupture_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_shear_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_bolt_shear_bot_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_bolt_shear_top_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_nodeform_bot_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_deform_bot_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_nodeform_top_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltshaft_bearing_deform_top_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_nodeform_bot_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_deform_bot_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_nodeform_top_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_boltreinf_bearing_deform_top_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_trans_bot_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_long_bot_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_trans_top_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_weld_long_top_h.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ind_default.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("reinf_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("pole_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_reinf_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("name")
        SQLInsertFields = SQLInsertFields.AddtoDBString("type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("b")
        SQLInsertFields = SQLInsertFields.AddtoDBString("h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sr_diam")
        SQLInsertFields = SQLInsertFields.AddtoDBString("channel_thkns_web")
        SQLInsertFields = SQLInsertFields.AddtoDBString("channel_thkns_flange")
        SQLInsertFields = SQLInsertFields.AddtoDBString("channel_eo")
        SQLInsertFields = SQLInsertFields.AddtoDBString("channel_J")
        SQLInsertFields = SQLInsertFields.AddtoDBString("channel_Cw")
        SQLInsertFields = SQLInsertFields.AddtoDBString("area_gross")
        SQLInsertFields = SQLInsertFields.AddtoDBString("centroid")
        SQLInsertFields = SQLInsertFields.AddtoDBString("istension")
        SQLInsertFields = SQLInsertFields.AddtoDBString("matl_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_matl_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Ix")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Iy")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Lu")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Kx")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Ky")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_hole_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("area_net")
        SQLInsertFields = SQLInsertFields.AddtoDBString("shear_lag")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_type_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_cap_revF_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_cap_revG_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_cap_revH_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_id_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_bolt_id_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_N_or_X_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_num_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_spacing_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_edge_dist_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FlangeOrBP_connected_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_grade_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_trans_type_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_trans_length_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_groove_depth_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_groove_angle_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_trans_fillet_size_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_trans_eff_throat_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_long_type_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_long_length_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_long_fillet_size_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_long_eff_throat_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("top_bot_connections_symmetrical")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_type_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_cap_revF_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_cap_revG_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_cap_revH_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_id_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_bolt_id_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_N_or_X_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_num_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_spacing_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_edge_dist_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("FlangeOrBP_connected_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_grade_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_trans_type_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_trans_length_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_groove_depth_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_groove_angle_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_trans_fillet_size_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_trans_eff_throat_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_long_type_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_long_length_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_long_fillet_size_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_long_eff_throat_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("conn_length_channel")
        SQLInsertFields = SQLInsertFields.AddtoDBString("conn_length_bot")
        SQLInsertFields = SQLInsertFields.AddtoDBString("conn_length_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_comp_xx_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_comp_yy_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_tens_yield_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_tens_rupture_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_shear_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_bolt_shear_bot_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_bolt_shear_top_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_nodeform_bot_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_deform_bot_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_nodeform_top_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_deform_top_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_nodeform_bot_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_deform_bot_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_nodeform_top_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_deform_top_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_trans_bot_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_long_bot_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_trans_top_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_long_top_f")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_comp_xx_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_comp_yy_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_tens_yield_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_tens_rupture_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_shear_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_bolt_shear_bot_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_bolt_shear_top_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_nodeform_bot_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_deform_bot_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_nodeform_top_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_deform_top_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_nodeform_bot_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_deform_bot_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_nodeform_top_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_deform_top_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_trans_bot_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_long_bot_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_trans_top_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_long_top_g")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_comp_xx_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_comp_yy_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_tens_yield_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_tens_rupture_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_shear_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_bolt_shear_bot_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_bolt_shear_top_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_nodeform_bot_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_deform_bot_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_nodeform_top_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltshaft_bearing_deform_top_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_nodeform_bot_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_deform_bot_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_nodeform_top_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_boltreinf_bearing_deform_top_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_trans_bot_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_long_bot_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_trans_top_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_weld_long_top_h")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ind_default")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("reinf_id = " & Me.reinf_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pole_id = " & Me.pole_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_reinf_id = " & Me.local_reinf_id.ToString.FormatDBValue) 'Commented out so when searching if a matching entry exists within DB already, the local ID wont disqualify it
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("name = " & Me.name.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("type = " & Me.type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("b = " & Me.b.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("h = " & Me.h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sr_diam = " & Me.sr_diam.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("channel_thkns_web = " & Me.channel_thkns_web.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("channel_thkns_flange = " & Me.channel_thkns_flange.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("channel_eo = " & Me.channel_eo.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("channel_J = " & Me.channel_J.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("channel_Cw = " & Me.channel_Cw.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("area_gross = " & Me.area_gross.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("centroid = " & Me.centroid.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("istension = " & Me.istension.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("matl_id = @SubLevel4ID") '& Me.matl_id.ToString.FormatDBValue) ''''''''''
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_matl_id = " & Me.local_matl_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Ix = " & Me.Ix.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Iy = " & Me.Iy.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Lu = " & Me.Lu.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Kx = " & Me.Kx.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Ky = " & Me.Ky.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_hole_size = " & Me.bolt_hole_size.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("area_net = " & Me.area_net.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("shear_lag = " & Me.shear_lag.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_type_bot = " & Me.connection_type_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_cap_revF_bot = " & Me.connection_cap_revF_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_cap_revG_bot = " & Me.connection_cap_revG_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_cap_revH_bot = " & Me.connection_cap_revH_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_id_bot = @BotBoltID") '& Me.bolt_id_bot.ToString.FormatDBValue) ''''''''''
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_bolt_id_bot = " & Me.local_bolt_id_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_N_or_X_bot = " & Me.bolt_N_or_X_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_num_bot = " & Me.bolt_num_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_spacing_bot = " & Me.bolt_spacing_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_edge_dist_bot = " & Me.bolt_edge_dist_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("FlangeOrBP_connected_bot = " & Me.FlangeOrBP_connected_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_grade_bot = " & Me.weld_grade_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_trans_type_bot = " & Me.weld_trans_type_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_trans_length_bot = " & Me.weld_trans_length_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_groove_depth_bot = " & Me.weld_groove_depth_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_groove_angle_bot = " & Me.weld_groove_angle_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_trans_fillet_size_bot = " & Me.weld_trans_fillet_size_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_trans_eff_throat_bot = " & Me.weld_trans_eff_throat_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_long_type_bot = " & Me.weld_long_type_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_long_length_bot = " & Me.weld_long_length_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_long_fillet_size_bot = " & Me.weld_long_fillet_size_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_long_eff_throat_bot = " & Me.weld_long_eff_throat_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("top_bot_connections_symmetrical = " & Me.top_bot_connections_symmetrical.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_type_top = " & Me.connection_type_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_cap_revF_top = " & Me.connection_cap_revF_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_cap_revG_top = " & Me.connection_cap_revG_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_cap_revH_top = " & Me.connection_cap_revH_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_id_top = @TopBoltID") '& Me.bolt_id_top.ToString.FormatDBValue) ''''''''''
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_bolt_id_top = " & Me.local_bolt_id_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_N_or_X_top = " & Me.bolt_N_or_X_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_num_top = " & Me.bolt_num_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_spacing_top = " & Me.bolt_spacing_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_edge_dist_top = " & Me.bolt_edge_dist_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("FlangeOrBP_connected_top = " & Me.FlangeOrBP_connected_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_grade_top = " & Me.weld_grade_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_trans_type_top = " & Me.weld_trans_type_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_trans_length_top = " & Me.weld_trans_length_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_groove_depth_top = " & Me.weld_groove_depth_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_groove_angle_top = " & Me.weld_groove_angle_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_trans_fillet_size_top = " & Me.weld_trans_fillet_size_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_trans_eff_throat_top = " & Me.weld_trans_eff_throat_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_long_type_top = " & Me.weld_long_type_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_long_length_top = " & Me.weld_long_length_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_long_fillet_size_top = " & Me.weld_long_fillet_size_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_long_eff_throat_top = " & Me.weld_long_eff_throat_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("conn_length_channel = " & Me.conn_length_channel.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("conn_length_bot = " & Me.conn_length_bot.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("conn_length_top = " & Me.conn_length_top.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_comp_xx_f = " & Me.cap_comp_xx_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_comp_yy_f = " & Me.cap_comp_yy_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_tens_yield_f = " & Me.cap_tens_yield_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_tens_rupture_f = " & Me.cap_tens_rupture_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_shear_f = " & Me.cap_shear_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_bolt_shear_bot_f = " & Me.cap_bolt_shear_bot_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_bolt_shear_top_f = " & Me.cap_bolt_shear_top_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_nodeform_bot_f = " & Me.cap_boltshaft_bearing_nodeform_bot_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_deform_bot_f = " & Me.cap_boltshaft_bearing_deform_bot_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_nodeform_top_f = " & Me.cap_boltshaft_bearing_nodeform_top_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_deform_top_f = " & Me.cap_boltshaft_bearing_deform_top_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_nodeform_bot_f = " & Me.cap_boltreinf_bearing_nodeform_bot_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_deform_bot_f = " & Me.cap_boltreinf_bearing_deform_bot_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_nodeform_top_f = " & Me.cap_boltreinf_bearing_nodeform_top_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_deform_top_f = " & Me.cap_boltreinf_bearing_deform_top_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_trans_bot_f = " & Me.cap_weld_trans_bot_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_long_bot_f = " & Me.cap_weld_long_bot_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_trans_top_f = " & Me.cap_weld_trans_top_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_long_top_f = " & Me.cap_weld_long_top_f.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_comp_xx_g = " & Me.cap_comp_xx_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_comp_yy_g = " & Me.cap_comp_yy_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_tens_yield_g = " & Me.cap_tens_yield_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_tens_rupture_g = " & Me.cap_tens_rupture_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_shear_g = " & Me.cap_shear_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_bolt_shear_bot_g = " & Me.cap_bolt_shear_bot_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_bolt_shear_top_g = " & Me.cap_bolt_shear_top_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_nodeform_bot_g = " & Me.cap_boltshaft_bearing_nodeform_bot_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_deform_bot_g = " & Me.cap_boltshaft_bearing_deform_bot_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_nodeform_top_g = " & Me.cap_boltshaft_bearing_nodeform_top_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_deform_top_g = " & Me.cap_boltshaft_bearing_deform_top_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_nodeform_bot_g = " & Me.cap_boltreinf_bearing_nodeform_bot_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_deform_bot_g = " & Me.cap_boltreinf_bearing_deform_bot_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_nodeform_top_g = " & Me.cap_boltreinf_bearing_nodeform_top_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_deform_top_g = " & Me.cap_boltreinf_bearing_deform_top_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_trans_bot_g = " & Me.cap_weld_trans_bot_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_long_bot_g = " & Me.cap_weld_long_bot_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_trans_top_g = " & Me.cap_weld_trans_top_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_long_top_g = " & Me.cap_weld_long_top_g.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_comp_xx_h = " & Me.cap_comp_xx_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_comp_yy_h = " & Me.cap_comp_yy_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_tens_yield_h = " & Me.cap_tens_yield_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_tens_rupture_h = " & Me.cap_tens_rupture_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_shear_h = " & Me.cap_shear_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_bolt_shear_bot_h = " & Me.cap_bolt_shear_bot_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_bolt_shear_top_h = " & Me.cap_bolt_shear_top_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_nodeform_bot_h = " & Me.cap_boltshaft_bearing_nodeform_bot_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_deform_bot_h = " & Me.cap_boltshaft_bearing_deform_bot_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_nodeform_top_h = " & Me.cap_boltshaft_bearing_nodeform_top_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltshaft_bearing_deform_top_h = " & Me.cap_boltshaft_bearing_deform_top_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_nodeform_bot_h = " & Me.cap_boltreinf_bearing_nodeform_bot_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_deform_bot_h = " & Me.cap_boltreinf_bearing_deform_bot_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_nodeform_top_h = " & Me.cap_boltreinf_bearing_nodeform_top_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_boltreinf_bearing_deform_top_h = " & Me.cap_boltreinf_bearing_deform_top_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_trans_bot_h = " & Me.cap_weld_trans_bot_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_long_bot_h = " & Me.cap_weld_long_bot_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_trans_top_h = " & Me.cap_weld_trans_top_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cap_weld_long_top_h = " & Me.cap_weld_long_top_h.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ind_default = " & Me.ind_default.ToString.FormatDBValue)

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
        Dim otherToCompare As PoleReinfProp = TryCast(other, PoleReinfProp)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.reinf_id.CheckChange(otherToCompare.reinf_id, changes, categoryName, "Reinf Id"), Equals, False)
        'Equals = If(Me.pole_id.CheckChange(otherToCompare.pole_id, changes, categoryName, "Pole Id"), Equals, False)
        Equals = If(Me.local_reinf_id.CheckChange(otherToCompare.local_reinf_id, changes, categoryName, "Local Reinf Id"), Equals, False)
        Equals = If(Me.name.CheckChange(otherToCompare.name, changes, categoryName, "Name"), Equals, False)
        Equals = If(Me.type.CheckChange(otherToCompare.type, changes, categoryName, "Type"), Equals, False)
        Equals = If(Me.b.CheckChange(otherToCompare.b, changes, categoryName, "B"), Equals, False)
        Equals = If(Me.h.CheckChange(otherToCompare.h, changes, categoryName, "H"), Equals, False)
        Equals = If(Me.sr_diam.CheckChange(otherToCompare.sr_diam, changes, categoryName, "Sr Diam"), Equals, False)
        Equals = If(Me.channel_thkns_web.CheckChange(otherToCompare.channel_thkns_web, changes, categoryName, "Channel Thkns Web"), Equals, False)
        Equals = If(Me.channel_thkns_flange.CheckChange(otherToCompare.channel_thkns_flange, changes, categoryName, "Channel Thkns Flange"), Equals, False)
        Equals = If(Me.channel_eo.CheckChange(otherToCompare.channel_eo, changes, categoryName, "Channel Eo"), Equals, False)
        Equals = If(Me.channel_J.CheckChange(otherToCompare.channel_J, changes, categoryName, "Channel J"), Equals, False)
        Equals = If(Me.channel_Cw.CheckChange(otherToCompare.channel_Cw, changes, categoryName, "Channel Cw"), Equals, False)
        Equals = If(Me.area_gross.CheckChange(otherToCompare.area_gross, changes, categoryName, "Area Gross"), Equals, False)
        Equals = If(Me.centroid.CheckChange(otherToCompare.centroid, changes, categoryName, "Centroid"), Equals, False)
        Equals = If(Me.istension.CheckChange(otherToCompare.istension, changes, categoryName, "Istension"), Equals, False)
        'Equals = If(Me.matl_id.CheckChange(otherToCompare.matl_id, changes, categoryName, "Matl Id"), Equals, False)
        Equals = If(Me.local_matl_id.CheckChange(otherToCompare.local_matl_id, changes, categoryName, "Local Matl Id"), Equals, False)
        Equals = If(Me.Ix.CheckChange(otherToCompare.Ix, changes, categoryName, "Ix"), Equals, False)
        Equals = If(Me.Iy.CheckChange(otherToCompare.Iy, changes, categoryName, "Iy"), Equals, False)
        Equals = If(Me.Lu.CheckChange(otherToCompare.Lu, changes, categoryName, "Lu"), Equals, False)
        Equals = If(Me.Kx.CheckChange(otherToCompare.Kx, changes, categoryName, "Kx"), Equals, False)
        Equals = If(Me.Ky.CheckChange(otherToCompare.Ky, changes, categoryName, "Ky"), Equals, False)
        Equals = If(Me.bolt_hole_size.CheckChange(otherToCompare.bolt_hole_size, changes, categoryName, "Bolt Hole Size"), Equals, False)
        Equals = If(Me.area_net.CheckChange(otherToCompare.area_net, changes, categoryName, "Area Net"), Equals, False)
        Equals = If(Me.shear_lag.CheckChange(otherToCompare.shear_lag, changes, categoryName, "Shear Lag"), Equals, False)
        Equals = If(Me.connection_type_bot.CheckChange(otherToCompare.connection_type_bot, changes, categoryName, "Connection Type Bot"), Equals, False)
        Equals = If(Me.connection_cap_revF_bot.CheckChange(otherToCompare.connection_cap_revF_bot, changes, categoryName, "Connection Cap Revf Bot"), Equals, False)
        Equals = If(Me.connection_cap_revG_bot.CheckChange(otherToCompare.connection_cap_revG_bot, changes, categoryName, "Connection Cap Revg Bot"), Equals, False)
        Equals = If(Me.connection_cap_revH_bot.CheckChange(otherToCompare.connection_cap_revH_bot, changes, categoryName, "Connection Cap Revh Bot"), Equals, False)
        'Equals = If(Me.bolt_id_bot.CheckChange(otherToCompare.bolt_id_bot, changes, categoryName, "Bolt Id Bot"), Equals, False)
        Equals = If(Me.local_bolt_id_bot.CheckChange(otherToCompare.local_bolt_id_bot, changes, categoryName, "Local Bolt Id Bot"), Equals, False)
        Equals = If(Me.bolt_N_or_X_bot.CheckChange(otherToCompare.bolt_N_or_X_bot, changes, categoryName, "Bolt N Or X Bot"), Equals, False)
        Equals = If(Me.bolt_num_bot.CheckChange(otherToCompare.bolt_num_bot, changes, categoryName, "Bolt Num Bot"), Equals, False)
        Equals = If(Me.bolt_spacing_bot.CheckChange(otherToCompare.bolt_spacing_bot, changes, categoryName, "Bolt Spacing Bot"), Equals, False)
        Equals = If(Me.bolt_edge_dist_bot.CheckChange(otherToCompare.bolt_edge_dist_bot, changes, categoryName, "Bolt Edge Dist Bot"), Equals, False)
        Equals = If(Me.FlangeOrBP_connected_bot.CheckChange(otherToCompare.FlangeOrBP_connected_bot, changes, categoryName, "Flangeorbp Connected Bot"), Equals, False)
        Equals = If(Me.weld_grade_bot.CheckChange(otherToCompare.weld_grade_bot, changes, categoryName, "Weld Grade Bot"), Equals, False)
        Equals = If(Me.weld_trans_type_bot.CheckChange(otherToCompare.weld_trans_type_bot, changes, categoryName, "Weld Trans Type Bot"), Equals, False)
        Equals = If(Me.weld_trans_length_bot.CheckChange(otherToCompare.weld_trans_length_bot, changes, categoryName, "Weld Trans Length Bot"), Equals, False)
        Equals = If(Me.weld_groove_depth_bot.CheckChange(otherToCompare.weld_groove_depth_bot, changes, categoryName, "Weld Groove Depth Bot"), Equals, False)
        Equals = If(Me.weld_groove_angle_bot.CheckChange(otherToCompare.weld_groove_angle_bot, changes, categoryName, "Weld Groove Angle Bot"), Equals, False)
        Equals = If(Me.weld_trans_fillet_size_bot.CheckChange(otherToCompare.weld_trans_fillet_size_bot, changes, categoryName, "Weld Trans Fillet Size Bot"), Equals, False)
        Equals = If(Me.weld_trans_eff_throat_bot.CheckChange(otherToCompare.weld_trans_eff_throat_bot, changes, categoryName, "Weld Trans Eff Throat Bot"), Equals, False)
        Equals = If(Me.weld_long_type_bot.CheckChange(otherToCompare.weld_long_type_bot, changes, categoryName, "Weld Long Type Bot"), Equals, False)
        Equals = If(Me.weld_long_length_bot.CheckChange(otherToCompare.weld_long_length_bot, changes, categoryName, "Weld Long Length Bot"), Equals, False)
        Equals = If(Me.weld_long_fillet_size_bot.CheckChange(otherToCompare.weld_long_fillet_size_bot, changes, categoryName, "Weld Long Fillet Size Bot"), Equals, False)
        Equals = If(Me.weld_long_eff_throat_bot.CheckChange(otherToCompare.weld_long_eff_throat_bot, changes, categoryName, "Weld Long Eff Throat Bot"), Equals, False)
        Equals = If(Me.top_bot_connections_symmetrical.CheckChange(otherToCompare.top_bot_connections_symmetrical, changes, categoryName, "Top Bot Connections Symmetrical"), Equals, False)
        Equals = If(Me.connection_type_top.CheckChange(otherToCompare.connection_type_top, changes, categoryName, "Connection Type Top"), Equals, False)
        Equals = If(Me.connection_cap_revF_top.CheckChange(otherToCompare.connection_cap_revF_top, changes, categoryName, "Connection Cap Revf Top"), Equals, False)
        Equals = If(Me.connection_cap_revG_top.CheckChange(otherToCompare.connection_cap_revG_top, changes, categoryName, "Connection Cap Revg Top"), Equals, False)
        Equals = If(Me.connection_cap_revH_top.CheckChange(otherToCompare.connection_cap_revH_top, changes, categoryName, "Connection Cap Revh Top"), Equals, False)
        'Equals = If(Me.bolt_id_top.CheckChange(otherToCompare.bolt_id_top, changes, categoryName, "Bolt Id Top"), Equals, False)
        Equals = If(Me.local_bolt_id_top.CheckChange(otherToCompare.local_bolt_id_top, changes, categoryName, "Local Bolt Id Top"), Equals, False)
        Equals = If(Me.bolt_N_or_X_top.CheckChange(otherToCompare.bolt_N_or_X_top, changes, categoryName, "Bolt N Or X Top"), Equals, False)
        Equals = If(Me.bolt_num_top.CheckChange(otherToCompare.bolt_num_top, changes, categoryName, "Bolt Num Top"), Equals, False)
        Equals = If(Me.bolt_spacing_top.CheckChange(otherToCompare.bolt_spacing_top, changes, categoryName, "Bolt Spacing Top"), Equals, False)
        Equals = If(Me.bolt_edge_dist_top.CheckChange(otherToCompare.bolt_edge_dist_top, changes, categoryName, "Bolt Edge Dist Top"), Equals, False)
        Equals = If(Me.FlangeOrBP_connected_top.CheckChange(otherToCompare.FlangeOrBP_connected_top, changes, categoryName, "Flangeorbp Connected Top"), Equals, False)
        Equals = If(Me.weld_grade_top.CheckChange(otherToCompare.weld_grade_top, changes, categoryName, "Weld Grade Top"), Equals, False)
        Equals = If(Me.weld_trans_type_top.CheckChange(otherToCompare.weld_trans_type_top, changes, categoryName, "Weld Trans Type Top"), Equals, False)
        Equals = If(Me.weld_trans_length_top.CheckChange(otherToCompare.weld_trans_length_top, changes, categoryName, "Weld Trans Length Top"), Equals, False)
        Equals = If(Me.weld_groove_depth_top.CheckChange(otherToCompare.weld_groove_depth_top, changes, categoryName, "Weld Groove Depth Top"), Equals, False)
        Equals = If(Me.weld_groove_angle_top.CheckChange(otherToCompare.weld_groove_angle_top, changes, categoryName, "Weld Groove Angle Top"), Equals, False)
        Equals = If(Me.weld_trans_fillet_size_top.CheckChange(otherToCompare.weld_trans_fillet_size_top, changes, categoryName, "Weld Trans Fillet Size Top"), Equals, False)
        Equals = If(Me.weld_trans_eff_throat_top.CheckChange(otherToCompare.weld_trans_eff_throat_top, changes, categoryName, "Weld Trans Eff Throat Top"), Equals, False)
        Equals = If(Me.weld_long_type_top.CheckChange(otherToCompare.weld_long_type_top, changes, categoryName, "Weld Long Type Top"), Equals, False)
        Equals = If(Me.weld_long_length_top.CheckChange(otherToCompare.weld_long_length_top, changes, categoryName, "Weld Long Length Top"), Equals, False)
        Equals = If(Me.weld_long_fillet_size_top.CheckChange(otherToCompare.weld_long_fillet_size_top, changes, categoryName, "Weld Long Fillet Size Top"), Equals, False)
        Equals = If(Me.weld_long_eff_throat_top.CheckChange(otherToCompare.weld_long_eff_throat_top, changes, categoryName, "Weld Long Eff Throat Top"), Equals, False)
        Equals = If(Me.conn_length_channel.CheckChange(otherToCompare.conn_length_channel, changes, categoryName, "Conn Length Channel"), Equals, False)
        Equals = If(Me.conn_length_bot.CheckChange(otherToCompare.conn_length_bot, changes, categoryName, "Conn Length Bot"), Equals, False)
        Equals = If(Me.conn_length_top.CheckChange(otherToCompare.conn_length_top, changes, categoryName, "Conn Length Top"), Equals, False)
        Equals = If(Me.cap_comp_xx_f.CheckChange(otherToCompare.cap_comp_xx_f, changes, categoryName, "Cap Comp Xx F"), Equals, False)
        Equals = If(Me.cap_comp_yy_f.CheckChange(otherToCompare.cap_comp_yy_f, changes, categoryName, "Cap Comp Yy F"), Equals, False)
        Equals = If(Me.cap_tens_yield_f.CheckChange(otherToCompare.cap_tens_yield_f, changes, categoryName, "Cap Tens Yield F"), Equals, False)
        Equals = If(Me.cap_tens_rupture_f.CheckChange(otherToCompare.cap_tens_rupture_f, changes, categoryName, "Cap Tens Rupture F"), Equals, False)
        Equals = If(Me.cap_shear_f.CheckChange(otherToCompare.cap_shear_f, changes, categoryName, "Cap Shear F"), Equals, False)
        Equals = If(Me.cap_bolt_shear_bot_f.CheckChange(otherToCompare.cap_bolt_shear_bot_f, changes, categoryName, "Cap Bolt Shear Bot F"), Equals, False)
        Equals = If(Me.cap_bolt_shear_top_f.CheckChange(otherToCompare.cap_bolt_shear_top_f, changes, categoryName, "Cap Bolt Shear Top F"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_nodeform_bot_f.CheckChange(otherToCompare.cap_boltshaft_bearing_nodeform_bot_f, changes, categoryName, "Cap Boltshaft Bearing Nodeform Bot F"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_deform_bot_f.CheckChange(otherToCompare.cap_boltshaft_bearing_deform_bot_f, changes, categoryName, "Cap Boltshaft Bearing Deform Bot F"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_nodeform_top_f.CheckChange(otherToCompare.cap_boltshaft_bearing_nodeform_top_f, changes, categoryName, "Cap Boltshaft Bearing Nodeform Top F"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_deform_top_f.CheckChange(otherToCompare.cap_boltshaft_bearing_deform_top_f, changes, categoryName, "Cap Boltshaft Bearing Deform Top F"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_nodeform_bot_f.CheckChange(otherToCompare.cap_boltreinf_bearing_nodeform_bot_f, changes, categoryName, "Cap Boltreinf Bearing Nodeform Bot F"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_deform_bot_f.CheckChange(otherToCompare.cap_boltreinf_bearing_deform_bot_f, changes, categoryName, "Cap Boltreinf Bearing Deform Bot F"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_nodeform_top_f.CheckChange(otherToCompare.cap_boltreinf_bearing_nodeform_top_f, changes, categoryName, "Cap Boltreinf Bearing Nodeform Top F"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_deform_top_f.CheckChange(otherToCompare.cap_boltreinf_bearing_deform_top_f, changes, categoryName, "Cap Boltreinf Bearing Deform Top F"), Equals, False)
        Equals = If(Me.cap_weld_trans_bot_f.CheckChange(otherToCompare.cap_weld_trans_bot_f, changes, categoryName, "Cap Weld Trans Bot F"), Equals, False)
        Equals = If(Me.cap_weld_long_bot_f.CheckChange(otherToCompare.cap_weld_long_bot_f, changes, categoryName, "Cap Weld Long Bot F"), Equals, False)
        Equals = If(Me.cap_weld_trans_top_f.CheckChange(otherToCompare.cap_weld_trans_top_f, changes, categoryName, "Cap Weld Trans Top F"), Equals, False)
        Equals = If(Me.cap_weld_long_top_f.CheckChange(otherToCompare.cap_weld_long_top_f, changes, categoryName, "Cap Weld Long Top F"), Equals, False)
        Equals = If(Me.cap_comp_xx_g.CheckChange(otherToCompare.cap_comp_xx_g, changes, categoryName, "Cap Comp Xx G"), Equals, False)
        Equals = If(Me.cap_comp_yy_g.CheckChange(otherToCompare.cap_comp_yy_g, changes, categoryName, "Cap Comp Yy G"), Equals, False)
        Equals = If(Me.cap_tens_yield_g.CheckChange(otherToCompare.cap_tens_yield_g, changes, categoryName, "Cap Tens Yield G"), Equals, False)
        Equals = If(Me.cap_tens_rupture_g.CheckChange(otherToCompare.cap_tens_rupture_g, changes, categoryName, "Cap Tens Rupture G"), Equals, False)
        Equals = If(Me.cap_shear_g.CheckChange(otherToCompare.cap_shear_g, changes, categoryName, "Cap Shear G"), Equals, False)
        Equals = If(Me.cap_bolt_shear_bot_g.CheckChange(otherToCompare.cap_bolt_shear_bot_g, changes, categoryName, "Cap Bolt Shear Bot G"), Equals, False)
        Equals = If(Me.cap_bolt_shear_top_g.CheckChange(otherToCompare.cap_bolt_shear_top_g, changes, categoryName, "Cap Bolt Shear Top G"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_nodeform_bot_g.CheckChange(otherToCompare.cap_boltshaft_bearing_nodeform_bot_g, changes, categoryName, "Cap Boltshaft Bearing Nodeform Bot G"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_deform_bot_g.CheckChange(otherToCompare.cap_boltshaft_bearing_deform_bot_g, changes, categoryName, "Cap Boltshaft Bearing Deform Bot G"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_nodeform_top_g.CheckChange(otherToCompare.cap_boltshaft_bearing_nodeform_top_g, changes, categoryName, "Cap Boltshaft Bearing Nodeform Top G"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_deform_top_g.CheckChange(otherToCompare.cap_boltshaft_bearing_deform_top_g, changes, categoryName, "Cap Boltshaft Bearing Deform Top G"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_nodeform_bot_g.CheckChange(otherToCompare.cap_boltreinf_bearing_nodeform_bot_g, changes, categoryName, "Cap Boltreinf Bearing Nodeform Bot G"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_deform_bot_g.CheckChange(otherToCompare.cap_boltreinf_bearing_deform_bot_g, changes, categoryName, "Cap Boltreinf Bearing Deform Bot G"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_nodeform_top_g.CheckChange(otherToCompare.cap_boltreinf_bearing_nodeform_top_g, changes, categoryName, "Cap Boltreinf Bearing Nodeform Top G"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_deform_top_g.CheckChange(otherToCompare.cap_boltreinf_bearing_deform_top_g, changes, categoryName, "Cap Boltreinf Bearing Deform Top G"), Equals, False)
        Equals = If(Me.cap_weld_trans_bot_g.CheckChange(otherToCompare.cap_weld_trans_bot_g, changes, categoryName, "Cap Weld Trans Bot G"), Equals, False)
        Equals = If(Me.cap_weld_long_bot_g.CheckChange(otherToCompare.cap_weld_long_bot_g, changes, categoryName, "Cap Weld Long Bot G"), Equals, False)
        Equals = If(Me.cap_weld_trans_top_g.CheckChange(otherToCompare.cap_weld_trans_top_g, changes, categoryName, "Cap Weld Trans Top G"), Equals, False)
        Equals = If(Me.cap_weld_long_top_g.CheckChange(otherToCompare.cap_weld_long_top_g, changes, categoryName, "Cap Weld Long Top G"), Equals, False)
        Equals = If(Me.cap_comp_xx_h.CheckChange(otherToCompare.cap_comp_xx_h, changes, categoryName, "Cap Comp Xx H"), Equals, False)
        Equals = If(Me.cap_comp_yy_h.CheckChange(otherToCompare.cap_comp_yy_h, changes, categoryName, "Cap Comp Yy H"), Equals, False)
        Equals = If(Me.cap_tens_yield_h.CheckChange(otherToCompare.cap_tens_yield_h, changes, categoryName, "Cap Tens Yield H"), Equals, False)
        Equals = If(Me.cap_tens_rupture_h.CheckChange(otherToCompare.cap_tens_rupture_h, changes, categoryName, "Cap Tens Rupture H"), Equals, False)
        Equals = If(Me.cap_shear_h.CheckChange(otherToCompare.cap_shear_h, changes, categoryName, "Cap Shear H"), Equals, False)
        Equals = If(Me.cap_bolt_shear_bot_h.CheckChange(otherToCompare.cap_bolt_shear_bot_h, changes, categoryName, "Cap Bolt Shear Bot H"), Equals, False)
        Equals = If(Me.cap_bolt_shear_top_h.CheckChange(otherToCompare.cap_bolt_shear_top_h, changes, categoryName, "Cap Bolt Shear Top H"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_nodeform_bot_h.CheckChange(otherToCompare.cap_boltshaft_bearing_nodeform_bot_h, changes, categoryName, "Cap Boltshaft Bearing Nodeform Bot H"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_deform_bot_h.CheckChange(otherToCompare.cap_boltshaft_bearing_deform_bot_h, changes, categoryName, "Cap Boltshaft Bearing Deform Bot H"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_nodeform_top_h.CheckChange(otherToCompare.cap_boltshaft_bearing_nodeform_top_h, changes, categoryName, "Cap Boltshaft Bearing Nodeform Top H"), Equals, False)
        Equals = If(Me.cap_boltshaft_bearing_deform_top_h.CheckChange(otherToCompare.cap_boltshaft_bearing_deform_top_h, changes, categoryName, "Cap Boltshaft Bearing Deform Top H"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_nodeform_bot_h.CheckChange(otherToCompare.cap_boltreinf_bearing_nodeform_bot_h, changes, categoryName, "Cap Boltreinf Bearing Nodeform Bot H"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_deform_bot_h.CheckChange(otherToCompare.cap_boltreinf_bearing_deform_bot_h, changes, categoryName, "Cap Boltreinf Bearing Deform Bot H"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_nodeform_top_h.CheckChange(otherToCompare.cap_boltreinf_bearing_nodeform_top_h, changes, categoryName, "Cap Boltreinf Bearing Nodeform Top H"), Equals, False)
        Equals = If(Me.cap_boltreinf_bearing_deform_top_h.CheckChange(otherToCompare.cap_boltreinf_bearing_deform_top_h, changes, categoryName, "Cap Boltreinf Bearing Deform Top H"), Equals, False)
        Equals = If(Me.cap_weld_trans_bot_h.CheckChange(otherToCompare.cap_weld_trans_bot_h, changes, categoryName, "Cap Weld Trans Bot H"), Equals, False)
        Equals = If(Me.cap_weld_long_bot_h.CheckChange(otherToCompare.cap_weld_long_bot_h, changes, categoryName, "Cap Weld Long Bot H"), Equals, False)
        Equals = If(Me.cap_weld_trans_top_h.CheckChange(otherToCompare.cap_weld_trans_top_h, changes, categoryName, "Cap Weld Trans Top H"), Equals, False)
        Equals = If(Me.cap_weld_long_top_h.CheckChange(otherToCompare.cap_weld_long_top_h, changes, categoryName, "Cap Weld Long Top H"), Equals, False)
        Equals = If(Me.ind_default.CheckChange(otherToCompare.ind_default, changes, categoryName, "Ind Default"), Equals, False)

    End Function
#End Region

End Class