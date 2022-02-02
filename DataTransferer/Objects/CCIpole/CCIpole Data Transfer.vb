Option Strict Off

Imports DevExpress.Spreadsheet
Imports System.Security.Principal

Partial Public Class DataTransfererCCIpole

#Region "Define"
    Private NewCCIpoleWb As New Workbook
    Private prop_ExcelFilePath As String

    Public Property Poles As New List(Of CCIpole)
    Public Property sqlPoles As New List(Of CCIpole)
    Private Property CCIpoleTemplatePath As String = "C:\Users\" & Environment.UserName & "\source\repos\Datatransferer NuGet\Reference\CCIpole (4.6.0) - TEMPLATE.xlsm"
    Private Property CCIpoleFileType As DocumentFormat = DocumentFormat.Xlsm

    Public Property poleDB As String
    Public Property poleID As WindowsIdentity

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
        poleID = LogOnUser
        poleDB = ActiveDatabase
        'BUNumber = BU 'Need to turn back on when connecting to dashboard. Turned off for testing.
        'STR_ID = Strucutre_ID 'Need to turn back on when connecting to dashboard. Turned off for testing.
    End Sub
#End Region

#Region "Load Data"
    Sub CreateSQLPoles(ByRef poleList As List(Of CCIpole))
        Dim refid As Integer
        Dim CCIpoleLoader As String

        'Load data to get CCIpole details for the existing structure model
        For Each item As SQLParameter In CCIpoleSQLDataTables()
            CCIpoleLoader = QueryBuilderFromFile(queryPath & "CCIpole\" & item.sqlQuery).Replace("[EXISTING MODEL]", GetExistingModelQuery())
            DoDaSQL.sqlLoader(CCIpoleLoader, item.sqlDatatable, ds, poleDB, poleID, "0")
            'If ds.Tables(item.sqlDatatable).Rows.Count = 0 Then Return False 'This may need adjusted since some tables can be empty
        Next

        'Custom Section to transfer data for the CCIpole tool. Needs to be adjusted for each tool.
        If IsSomething(ds.Tables("CCIpole Generale SQL")) Then
            For Each CCIpoleDataRow As DataRow In ds.Tables("CCIpole General SQL").Rows 'Help...No Details Table, main object is just the pole_structure_id and criteria_id
                refid = CType(CCIpoleDataRow.Item("pole_structure_id"), Integer)
                poleList.Add(New CCIpole(CCIpoleDataRow, refid))
            Next
        End If
    End Sub



    Public Function LoadFromEDS() As Boolean
        CreateSQLPoles(Poles)

        Return True
    End Function 'Create CCIpole objects based on what is saved in EDS

    Public Sub LoadFromExcel()
        'Dim refID As Integer
        'Dim refCol As String

        For Each item As EXCELDTParameter In CCIpoleExcelDTParameters()
            'Get tables from excel file 
            ds.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
        Next

        'Custom Section to transfer data for the CCIpole tool. Needs to be adjusted for each tool.
        'For Each CCIpoleDataRow As DataRow In ds.Tables("CCIpole General EXCEL").Rows
        '    refCol = "pole_structure_id"
        '    refID = CType(CCIpoleDataRow.Item(refCol), Integer)

        '    Poles.Add(New CCIpole(CCIpoleDataRow, refID))
        'Next

        Poles.Add(New CCIpole(ExcelFilePath))

        'Pull SQL data, if applicable, to compare with excel data
        CreateSQLPoles(sqlPoles)

        'If sqlPiles.Count > 0 Then 'same as if checking for id in tool, if ID greater than 0.
        For Each pole As CCIpole In Poles
            Dim IDmatch As Boolean = False
            If pole.pole_structure_id > 0 Then 'can skip loading SQL data if id = 0 (first time adding to EDS)
                For Each sqlpole As CCIpole In sqlPoles
                    If pole.pole_structure_id = sqlpole.pole_structure_id Then
                        IDmatch = True
                        If CheckChanges(pole, sqlpole) Then
                            isModelNeeded = True
                            isPoleNeeded = True
                        End If
                        Exit For
                    End If
                Next
                'IF ID match = False, Save the data because nothing exists in sql (could have copied tool from a different BU)
                If IDmatch = False Then
                    isModelNeeded = True
                    isPoleNeeded = True
                End If
            Else
                'Save the data because nothing exists in sql
                isModelNeeded = True
                isPoleNeeded = True
            End If
        Next

    End Sub 'Create CCIpole objects based on what is coming from the excel file
#End Region

#Region "Save Data"

    Sub Save1Pole(ByVal cp As CCIpole)

        Dim firstOne As Boolean = True
        Dim myCriteria As String = ""
        Dim myPoleSection As String = ""
        Dim myPoleReinfSection As String = ""
        Dim myReinfGroup As String = ""
        Dim myReinfDetail As String = ""
        Dim myIntGroup As String = ""
        Dim myIntDetail As String = ""
        Dim myReinfResults As String = ""
        Dim myReinfProp As String = ""
        Dim myBoltProp As String = ""
        Dim myMatlProp As String = ""

        Dim CCIpoleSaver As String = QueryBuilderFromFile(queryPath & "CCIpole\CCIpole (IN_UP).sql")
        CCIpoleSaver = CCIpoleSaver.Replace("[BU NUMBER]", BUNumber)
        CCIpoleSaver = CCIpoleSaver.Replace("[STRUCTURE ID]", STR_ID)
        'CCIpoleSaver = CCIpoleSaver.Replace("[FOUNDATION TYPE]", "Pile")

        If cp.pole_structure_id = 0 Or IsDBNull(cp.pole_structure_id) Then
            CCIpoleSaver = CCIpoleSaver.Replace("'[CCIPOLE ID]'", "NULL")
        Else
            CCIpoleSaver = CCIpoleSaver.Replace("[CCIPOLE ID]", cp.pole_structure_id.ToString)
        End If

        'Determine if new model ID needs created. Shouldn't be added to all individual tools (only needs to be referenced once)
        If isModelNeeded Then
            CCIpoleSaver = CCIpoleSaver.Replace("'[Model ID Needed]'", 1)
        Else
            CCIpoleSaver = CCIpoleSaver.Replace("'[Model ID Needed]'", 0)
        End If

        'Determine if new Pole ID needs created
        If isPoleNeeded Then
            CCIpoleSaver = CCIpoleSaver.Replace("'[CCIpole ID Needed]'", 1)
        Else
            CCIpoleSaver = CCIpoleSaver.Replace("'[CCIpole ID Needed]'", 0)
        End If

        'CCIpoleSaver = CCIpoleSaver.Replace("[INSERT ALL CCIPOLE DETAILS]", InsertPoleDetail(cp))
        'CCIpoleSaver = CCIpoleSaver.Replace("[INSERT ANALYSIS CRITERIA]", InsertPoleCriteria(cp.criteria))

        For Each cppc As PoleCriteria In cp.criteria
            If Not IsNothing(cppc.upper_structure_type) Or Not IsNothing(cppc.analysis_deg) Or Not IsNothing(cppc.geom_increment_length) Or Not IsNothing(cppc.vnum) Or Not IsNothing(cppc.check_connections) Or Not IsNothing(cppc.hole_deformation) Or Not IsNothing(cppc.ineff_mod_check) Or Not IsNothing(cppc.modified) Then
                Dim tempPoleCriteria As String = InsertPoleCriteria(cppc)
                If Not firstOne Then
                    myCriteria += ",(" & tempPoleCriteria & ")"
                Else
                    myCriteria += "(" & tempPoleCriteria & ")"
                End If
            End If
            firstOne = False
        Next
        firstOne = True
        CCIpoleSaver = CCIpoleSaver.Replace("'[INSERT POLE CRITERIA]'", myCriteria)

        For Each cpps As PoleSection In cp.unreinf_sections
            If Not IsNothing(cpps.local_section_id) Or Not IsNothing(cpps.elev_bot) Or Not IsNothing(cpps.elev_top) Or Not IsNothing(cpps.length_section) Or Not IsNothing(cpps.length_splice) Or Not IsNothing(cpps.num_sides) Or Not IsNothing(cpps.diam_bot) Or Not IsNothing(cpps.diam_top) Or Not IsNothing(cpps.wall_thickness) Or Not IsNothing(cpps.bend_radius) Or Not IsNothing(cpps.steel_grade_id) Or Not IsNothing(cpps.pole_type) Or Not IsNothing(cpps.section_name) Or Not IsNothing(cpps.socket_length) Or Not IsNothing(cpps.weight_mult) Or Not IsNothing(cpps.wp_mult) Or Not IsNothing(cpps.af_factor) Or Not IsNothing(cpps.ar_factor) Or Not IsNothing(cpps.round_area_ratio) Or Not IsNothing(cpps.flat_area_ratio) Then

                Dim tempPoleSection As String = InsertPoleSection(cpps)
                'If Not firstOne Then
                '    myPoleSection += ",(" & tempPoleSection & ")"
                'Else
                '    myPoleSection += "(" & tempPoleSection & ")"
                'End If
                myPoleSection = tempPoleSection

                Dim SubQuery_BGeom As String = QueryBuilderFromFile(queryPath & "CCIpole\CCIpole (SubQuery_GeomBase).sql")
                SubQuery_BGeom = SubQuery_BGeom.Replace("'[INSERT POLE SECTION]'", myPoleSection)

                For Each cpm As PropMatl In cp.matls
                    If cpm.matl_id = cpps.local_matl_id Then
                        Dim tempPropMatl As String = InsertPropMatl(cpm)
                        myMatlProp = tempPropMatl
                        SubQuery_BGeom = SubQuery_BGeom.Replace("'[INSERT MATL PROP]'", myMatlProp)
                    End If
                Next

                'firstOne = False

                CCIpoleSaver = CCIpoleSaver.Replace("--'[SUBQUERY]'", SubQuery_BGeom)
            End If
        Next
        'firstOne = True


        For Each cprs As PoleReinfSection In cp.reinf_sections
            Dim SubQuery_RGeom As String = QueryBuilderFromFile(queryPath & "CCIpole\CCIpole (IN_UP SUBQUERY R_Geom).sql")
            If Not IsNothing(cprs.local_section_id) Or Not IsNothing(cprs.elev_bot) Or Not IsNothing(cprs.elev_top) Or Not IsNothing(cprs.length_section) Or Not IsNothing(cprs.length_splice) Or Not IsNothing(cprs.num_sides) Or Not IsNothing(cprs.diam_bot) Or Not IsNothing(cprs.diam_top) Or Not IsNothing(cprs.wall_thickness) Or Not IsNothing(cprs.bend_radius) Or Not IsNothing(cprs.steel_grade_id) Or Not IsNothing(cprs.pole_type) Or Not IsNothing(cprs.weight_mult) Or Not IsNothing(cprs.section_name) Or Not IsNothing(cprs.socket_length) Or Not IsNothing(cprs.wp_mult) Or Not IsNothing(cprs.af_factor) Or Not IsNothing(cprs.ar_factor) Or Not IsNothing(cprs.round_area_ratio) Or Not IsNothing(cprs.flat_area_ratio) Then
                Dim tempPoleReinfSection As String = InsertPoleReinfSection(cprs)
                If Not firstOne Then
                    myPoleReinfSection += ",(" & tempPoleReinfSection & ")"
                Else
                    myPoleReinfSection += "(" & tempPoleReinfSection & ")"
                End If
            End If
            firstOne = False
            SubQuery_RGeom = SubQuery_RGeom.Replace("'[REINF SECTION]'", myPoleReinfSection)
            CCIpoleSaver = CCIpoleSaver.Replace("--'[SUBQUERY]'", SubQuery_RGeom)
        Next
        firstOne = True

        'For Each cprg As PoleReinfGroup In cp.reinf_groups
        '    If Not IsNothing(cprg.elev_bot_actual) Or Not IsNothing(cprg.elev_bot_eff) Or Not IsNothing(cprg.elev_top_actual) Or Not IsNothing(cprg.elev_top_eff) Or Not IsNothing(cprg.reinf_db_id) Then
        '        Dim tempPoleReinfGroup As String = InsertPoleReinfGroup(cprg)
        '        If Not firstOne Then
        '            myReinfGroup += ",(" & tempPoleReinfGroup & ")"
        '        Else
        '            myReinfGroup += "(" & tempPoleReinfGroup & ")"
        '        End If
        '    End If

        '    For Each cprd As PoleReinfDetail In cprg.reinf_ids
        '        If Not IsNothing(cprd.pole_flat) Or Not IsNothing(cprd.horizontal_offset) Or Not IsNothing(cprd.rotation) Or Not IsNothing(cprd.note) Then
        '            Dim tempPoleReinfDetail As String = InsertPoleReinfDetail(cprd)
        '            If Not firstOne Then
        '                myReinfDetail += ",(" & tempPoleReinfDetail & ")"
        '            Else
        '                myReinfDetail += "(" & tempPoleReinfDetail & ")"
        '            End If
        '        End If
        '        firstOne = False
        '    Next
        'Next
        'firstOne = True
        'CCIpoleSaver = CCIpoleSaver.Replace("([INSERT REINF GROUPS])", myReinfGroup)
        'CCIpoleSaver = CCIpoleSaver.Replace("([INSERT REINF DETAILS])", myReinfDetail)

        For Each cprg As PoleReinfGroup In cp.reinf_groups
            If Not IsNothing(cprg.elev_bot_actual) Or Not IsNothing(cprg.elev_bot_eff) Or Not IsNothing(cprg.elev_top_actual) Or Not IsNothing(cprg.elev_top_eff) Or Not IsNothing(cprg.reinf_db_id) Then

                Dim tempPoleReinfGroup As String = InsertPoleReinfGroup(cprg)
                myReinfGroup = tempPoleReinfGroup

                Dim subQuery_RGrp As String = QueryBuilderFromFile(queryPath & "CCIpole\CCIpole (IN_UP SUBQUERY Reinf Groups).sql")
                subQuery_RGrp = subQuery_RGrp.Replace("'[REINF GROUP]'", myReinfGroup)

                For Each cprd As PoleReinfDetail In cprg.reinf_ids
                    If cprd.local_group_id = cprg.local_group_id Then
                        If Not IsNothing(cprd.pole_flat) Or Not IsNothing(cprd.horizontal_offset) Or Not IsNothing(cprd.rotation) Or Not IsNothing(cprd.note) Then
                            Dim tempPoleReinfDetail As String = InsertPoleReinfDetail(cprd)
                            myReinfDetail = tempPoleReinfDetail
                        End If
                        subQuery_RGrp = subQuery_RGrp.Replace("'[REINF DETAILS]'", myReinfDetail)
                    End If
                    CCIpoleSaver = CCIpoleSaver.Replace("--'[SUBQUERY]'", subQuery_RGrp)
                Next

            End If

        Next


        For Each cpig As PoleIntGroup In cp.int_groups
            If Not IsNothing(cpig.elev_bot) Or Not IsNothing(cpig.elev_top) Or Not IsNothing(cpig.width) Or Not IsNothing(cpig.description) Then
                Dim tempPoleIntGroup As String = InsertPoleIntGroup(cpig)
                If Not firstOne Then
                    myIntGroup += ",(" & tempPoleIntGroup & ")"
                Else
                    myIntGroup += "(" & tempPoleIntGroup & ")"
                End If

                For Each cpid As PoleIntDetail In cpig.int_ids
                    If Not IsNothing(cpid.pole_flat) Or Not IsNothing(cpid.horizontal_offset) Or Not IsNothing(cpid.rotation) Or Not IsNothing(cpid.note) Then
                        Dim tempPoleIntDetail As String = InsertPoleIntDetail(cpid)
                        If Not firstOne Then
                            myIntDetail += ",(" & tempPoleIntDetail & ")"
                        Else
                            myIntDetail += "(" & tempPoleIntDetail & ")"
                        End If
                    End If
                    firstOne = False
                Next
            End If
        Next
        firstOne = True
        CCIpoleSaver = CCIpoleSaver.Replace("'[INSERT INT GROUPS]'", myIntGroup)
        CCIpoleSaver = CCIpoleSaver.Replace("'[INSERT INT DETAILS]'", myIntDetail)

        'For Each cpid As PoleIntDetail In cp.int_ids
        '    If Not IsNothing(cpid.pole_flat) Or Not IsNothing(cpid.horizontal_offset) Or Not IsNothing(cpid.rotation) Or Not IsNothing(cpid.note) Then
        '        Dim tempPoleIntDetail As String = InsertPoleIntDetail(cpid)
        '        If Not firstOne Then
        '            myIntDetail += ",(" & tempPoleIntDetail & ")"
        '        Else
        '            myIntDetail += "(" & tempPoleIntDetail & ")"
        '        End If
        '    End If
        '    firstOne = False
        'Next
        'firstOne = True
        'CCIpoleSaver = CCIpoleSaver.Replace("([INSERT INT DETAILS])", myIntDetail)

        For Each cprr As PoleReinfResults In cp.reinf_section_results
            If Not IsNothing(cprr.work_order_seq_num) Or Not IsNothing(cprr.reinf_group_id) Or Not IsNothing(cprr.result_lkup_value) Or Not IsNothing(cprr.rating) Then
                Dim tempPoleReinfResults As String = InsertPoleReinfResults(cprr)
                If Not firstOne Then
                    myReinfResults += ",(" & tempPoleReinfResults & ")"
                Else
                    myReinfResults += "(" & tempPoleReinfResults & ")"
                End If
            End If
            firstOne = False
        Next
        firstOne = True
        CCIpoleSaver = CCIpoleSaver.Replace("'[INSERT REINF RESULTS]'", myReinfResults)

        For Each cpr As PropReinf In cp.reinfs
            If Not IsNothing(cpr.name) Or Not IsNothing(cpr.type) Or Not IsNothing(cpr.b) Or Not IsNothing(cpr.h) Or Not IsNothing(cpr.sr_diam) Or Not IsNothing(cpr.channel_thkns_web) Or Not IsNothing(cpr.channel_thkns_flange) Or Not IsNothing(cpr.channel_eo) Or Not IsNothing(cpr.channel_J) Or Not IsNothing(cpr.channel_Cw) Or Not IsNothing(cpr.area_gross) Or Not IsNothing(cpr.centroid) Or Not IsNothing(cpr.istension) Or Not IsNothing(cpr.matl_id) Or Not IsNothing(cpr.Ix) Or Not IsNothing(cpr.Iy) Or Not IsNothing(cpr.Lu) Or Not IsNothing(cpr.Kx) Or Not IsNothing(cpr.Ky) Or Not IsNothing(cpr.bolt_hole_size) Or Not IsNothing(cpr.area_net) Or Not IsNothing(cpr.shear_lag) Or Not IsNothing(cpr.connection_type_bot) Or Not IsNothing(cpr.connection_cap_revF_bot) Or Not IsNothing(cpr.connection_cap_revG_bot) Or Not IsNothing(cpr.connection_cap_revH_bot) Or Not IsNothing(cpr.bolt_id_bot) Or Not IsNothing(cpr.bolt_N_or_X_bot) Or Not IsNothing(cpr.bolt_num_bot) Or Not IsNothing(cpr.bolt_spacing_bot) Or Not IsNothing(cpr.bolt_edge_dist_bot) Or Not IsNothing(cpr.FlangeOrBP_connected_bot) Or Not IsNothing(cpr.weld_grade_bot) Or Not IsNothing(cpr.weld_trans_type_bot) Or Not IsNothing(cpr.weld_trans_length_bot) Or Not IsNothing(cpr.weld_groove_depth_bot) Or Not IsNothing(cpr.weld_groove_angle_bot) Or Not IsNothing(cpr.weld_trans_fillet_size_bot) Or Not IsNothing(cpr.weld_trans_eff_throat_bot) Or Not IsNothing(cpr.weld_long_type_bot) Or Not IsNothing(cpr.weld_long_length_bot) Or Not IsNothing(cpr.weld_long_fillet_size_bot) Or Not IsNothing(cpr.weld_long_eff_throat_bot) Or Not IsNothing(cpr.top_bot_connections_symmetrical) Or Not IsNothing(cpr.connection_type_top) Or Not IsNothing(cpr.connection_cap_revF_top) Or Not IsNothing(cpr.connection_cap_revG_top) Or Not IsNothing(cpr.connection_cap_revH_top) Or Not IsNothing(cpr.bolt_id_top) Or Not IsNothing(cpr.bolt_N_or_X_top) Or Not IsNothing(cpr.bolt_num_top) Or Not IsNothing(cpr.bolt_spacing_top) Or Not IsNothing(cpr.bolt_edge_dist_top) Or Not IsNothing(cpr.FlangeOrBP_connected_top) Or Not IsNothing(cpr.weld_grade_top) Or Not IsNothing(cpr.weld_trans_type_top) Or Not IsNothing(cpr.weld_trans_length_top) Or Not IsNothing(cpr.weld_groove_depth_top) Or Not IsNothing(cpr.weld_groove_angle_top) Or Not IsNothing(cpr.weld_trans_fillet_size_top) Or Not IsNothing(cpr.weld_trans_eff_throat_top) Or Not IsNothing(cpr.weld_long_type_top) Or Not IsNothing(cpr.weld_long_length_top) Or Not IsNothing(cpr.weld_long_fillet_size_top) Or Not IsNothing(cpr.weld_long_eff_throat_top) Or Not IsNothing(cpr.conn_length_bot) Or Not IsNothing(cpr.conn_length_top) Or Not IsNothing(cpr.cap_comp_xx_f) Or Not IsNothing(cpr.cap_comp_yy_f) Or Not IsNothing(cpr.cap_tens_yield_f) Or Not IsNothing(cpr.cap_tens_rupture_f) Or Not IsNothing(cpr.cap_shear_f) Or Not IsNothing(cpr.cap_bolt_shear_bot_f) Or Not IsNothing(cpr.cap_bolt_shear_top_f) Or Not IsNothing(cpr.cap_boltshaft_bearing_nodeform_bot_f) Or Not IsNothing(cpr.cap_boltshaft_bearing_deform_bot_f) Or Not IsNothing(cpr.cap_boltshaft_bearing_nodeform_top_f) Or Not IsNothing(cpr.cap_boltshaft_bearing_deform_top_f) Or Not IsNothing(cpr.cap_boltreinf_bearing_nodeform_bot_f) Or Not IsNothing(cpr.cap_boltreinf_bearing_deform_bot_f) Or Not IsNothing(cpr.cap_boltreinf_bearing_nodeform_top_f) Or Not IsNothing(cpr.cap_boltreinf_bearing_deform_top_f) Or Not IsNothing(cpr.cap_weld_trans_bot_f) Or Not IsNothing(cpr.cap_weld_long_bot_f) Or Not IsNothing(cpr.cap_weld_trans_top_f) Or Not IsNothing(cpr.cap_weld_long_top_f) Or Not IsNothing(cpr.cap_comp_xx_g) Or Not IsNothing(cpr.cap_comp_yy_g) Or Not IsNothing(cpr.cap_tens_yield_g) Or Not IsNothing(cpr.cap_tens_rupture_g) Or Not IsNothing(cpr.cap_shear_g) Or Not IsNothing(cpr.cap_bolt_shear_bot_g) Or Not IsNothing(cpr.cap_bolt_shear_top_g) Or Not IsNothing(cpr.cap_boltshaft_bearing_deform_bot_g) Or Not IsNothing(cpr.cap_boltshaft_bearing_nodeform_top_g) Or Not IsNothing(cpr.cap_boltshaft_bearing_deform_top_g) Or Not IsNothing(cpr.cap_boltreinf_bearing_nodeform_bot_g) Or Not IsNothing(cpr.cap_boltreinf_bearing_deform_bot_g) Or Not IsNothing(cpr.cap_boltreinf_bearing_nodeform_top_g) Or Not IsNothing(cpr.cap_boltreinf_bearing_deform_top_g) Or Not IsNothing(cpr.cap_weld_trans_bot_g) Or Not IsNothing(cpr.cap_weld_long_bot_g) Or Not IsNothing(cpr.cap_weld_trans_top_g) Or Not IsNothing(cpr.cap_weld_long_top_g) Or Not IsNothing(cpr.cap_comp_xx_h) Or Not IsNothing(cpr.cap_comp_yy_h) Or Not IsNothing(cpr.cap_tens_yield_h) Or Not IsNothing(cpr.cap_tens_rupture_h) Or Not IsNothing(cpr.cap_shear_h) Or Not IsNothing(cpr.cap_bolt_shear_bot_h) Or Not IsNothing(cpr.cap_bolt_shear_top_h) Or Not IsNothing(cpr.cap_boltshaft_bearing_nodeform_bot_h) Or Not IsNothing(cpr.cap_boltshaft_bearing_deform_bot_h) Or Not IsNothing(cpr.cap_boltshaft_bearing_nodeform_top_h) Or Not IsNothing(cpr.cap_boltshaft_bearing_deform_top_h) Or Not IsNothing(cpr.cap_boltreinf_bearing_nodeform_bot_h) Or Not IsNothing(cpr.cap_boltreinf_bearing_deform_bot_h) Or Not IsNothing(cpr.cap_boltreinf_bearing_nodeform_top_h) Or Not IsNothing(cpr.cap_boltreinf_bearing_deform_top_h) Or Not IsNothing(cpr.cap_weld_trans_bot_h) Or Not IsNothing(cpr.cap_weld_long_bot_h) Or Not IsNothing(cpr.cap_weld_trans_top_h) Or Not IsNothing(cpr.cap_weld_long_top_h) Then
                Dim tempPropReinf As String = InsertPropReinf(cpr)
                If Not firstOne Then
                    myReinfProp += ",(" & tempPropReinf & ")"
                Else
                    myReinfProp += "(" & tempPropReinf & ")"
                End If
            End If
            firstOne = False
        Next
        firstOne = True
        CCIpoleSaver = CCIpoleSaver.Replace("'[INSERT ALL REINF PROP]'", myReinfProp)

        For Each cpb As PropBolt In cp.bolts
            If Not IsNothing(cpb.name) Or Not IsNothing(cpb.description) Or Not IsNothing(cpb.diam) Or Not IsNothing(cpb.area) Or Not IsNothing(cpb.fu_bolt) Or Not IsNothing(cpb.sleeve_diam_out) Or Not IsNothing(cpb.sleeve_diam_in) Or Not IsNothing(cpb.fu_sleeve) Or Not IsNothing(cpb.bolt_n_sleeve_shear_revF) Or Not IsNothing(cpb.bolt_x_sleeve_shear_revF) Or Not IsNothing(cpb.bolt_n_sleeve_shear_revG) Or Not IsNothing(cpb.bolt_x_sleeve_shear_revG) Or Not IsNothing(cpb.bolt_n_sleeve_shear_revH) Or Not IsNothing(cpb.bolt_x_sleeve_shear_revH) Or Not IsNothing(cpb.rb_applied_revH) Then
                Dim tempPropBolt As String = InsertPropBolt(cpb)
                If Not firstOne Then
                    myBoltProp += ",(" & tempPropBolt & ")"
                Else
                    myBoltProp += "(" & tempPropBolt & ")"
                End If
            End If
            firstOne = False
        Next
        firstOne = True
        CCIpoleSaver = CCIpoleSaver.Replace("'[INSERT ALL BOLT PROP')", myBoltProp)

        For Each cpm As PropMatl In cp.matls
            If Not IsNothing(cpm.name) Or Not IsNothing(cpm.fy) Or Not IsNothing(cpm.fu) Then
                Dim tempPropMatl As String = InsertPropMatl(cpm)
                If Not firstOne Then
                    myMatlProp += ",(" & tempPropMatl & ")"
                Else
                    myMatlProp += "(" & tempPropMatl & ")"
                End If
            End If
            firstOne = False
        Next
        firstOne = True
        CCIpoleSaver = CCIpoleSaver.Replace("'[INSERT MATL PROP]'", myMatlProp)

        myCriteria = ""
        myPoleSection = ""
        myPoleReinfSection = ""
        myReinfGroup = ""
        myReinfDetail = ""
        myIntGroup = ""
        myIntDetail = ""
        myReinfResults = ""
        myReinfProp = ""
        myBoltProp = ""
        myMatlProp = ""

        sqlSender(CCIPoleSaver, poleDB, poleID, "0")
    End Sub

    Public Sub SaveToEDS()
        For Each cp As CCIpole In Poles
            Save1Pole(cp)
        Next
    End Sub

    Public Sub SaveToExcel()

        For Each cp As CCIpole In Poles
            Dim row As Integer
            Dim col As Integer
            Dim drow, dcol As Integer

            LoadNewPole()
            With NewCCIpoleWb

                'pole_structure_id
                If Not IsNothing(cp.pole_structure_id) Then .Worksheets("Analysis Criteria (SAPI)").Range("A3").Value = CType(cp.pole_structure_id, Integer)

                'Analysis Criteria
                For Each pc As PoleCriteria In cp.criteria

                    col = 2

                    If Not IsNothing(pc.criteria_id) Then .Worksheets("Analysis Criteria (SAPI)").Cells(row, col).Value = CType(pc.criteria_id, Integer)
                    col += 1
                    If Not IsNothing(pc.upper_structure_type) Then .Worksheets("Analysis Criteria (SAPI)").Cells(row, col).Value = pc.upper_structure_type
                    col += 1
                    If Not IsNothing(pc.analysis_deg) Then .Worksheets("Analysis Criteria (SAPI)").Cells(row, col).Value = CType(pc.analysis_deg, Double)
                    col += 1
                    If Not IsNothing(pc.geom_increment_length) Then .Worksheets("Analysis Criteria (SAPI)").Cells(row, col).Value = CType(pc.geom_increment_length, Double)
                    col += 1
                    If Not IsNothing(pc.vnum) Then .Worksheets("Analysis Criteria (SAPI)").Cells(row, col).Value = pc.vnum
                    col += 1
                    If Not IsNothing(pc.check_connections) Then .Worksheets("Analysis Criteria (SAPI)").Cells(row, col).Value = pc.check_connections
                    col += 1
                    If Not IsNothing(pc.hole_deformation) Then .Worksheets("Analysis Criteria (SAPI)").Cells(row, col).Value = pc.hole_deformation
                    col += 1
                    If Not IsNothing(pc.ineff_mod_check) Then .Worksheets("Analysis Criteria (SAPI)").Cells(row, col).Value = pc.ineff_mod_check
                    col += 1
                    If Not IsNothing(pc.modified) Then .Worksheets("Analysis Criteria (SAPI)").Cells(row, col).Value = pc.modified

                    row += 1

                Next

                row = 3

                'Unreinforced Pole Sections
                For Each ps As PoleSection In cp.unreinf_sections

                    col = 1

                    If Not IsNothing(cp.pole_structure_id) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(cp.pole_structure_id, Integer)
                    col += 1
                    If Not IsNothing(ps.section_id) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.section_id, Integer)
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
                    If Not IsNothing(ps.bend_radius) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.bend_radius, Double)
                    col += 1
                    If Not IsNothing(ps.steel_grade_id) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = CType(ps.steel_grade_id, Integer)
                    col += 1
                    'If Not IsNothing(ps.pole_type) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = ps.pole_type
                    col += 1
                    If Not IsNothing(ps.section_name) Then .Worksheets("Unreinf Pole (SAPI)").Cells(row, col).Value = ps.section_name
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

                row = 3

                'Reinforced Pole Sections
                For Each prs As PoleReinfSection In cp.reinf_sections

                    col = 1

                    If Not IsNothing(cp.pole_structure_id) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(cp.pole_structure_id, Integer)
                    col += 1
                    If Not IsNothing(prs.section_id) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.section_id, Integer)
                    col += 1
                    If Not IsNothing(prs.local_section_id) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.local_section_id, Integer)
                    col += 1
                    If Not IsNothing(prs.elev_bot) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.elev_bot, Double)
                    col += 1
                    If Not IsNothing(prs.elev_top) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.elev_top, Double)
                    col += 1
                    If Not IsNothing(prs.length_section) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.length_section, Double)
                    col += 1
                    If Not IsNothing(prs.length_splice) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.length_splice, Double)
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
                    If Not IsNothing(prs.steel_grade_id) Then .Worksheets("Reinf Pole (SAPI)").Cells(row, col).Value = CType(prs.steel_grade_id, Integer)
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

                row = 3

                'Reinforcement Groups
                For Each prg As PoleReinfGroup In cp.reinf_groups

                    col = 1

                    If Not IsNothing(cp.pole_structure_id) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(cp.pole_structure_id, Integer)
                    col += 1
                    If Not IsNothing(prg.reinf_group_id) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.reinf_group_id, Integer)
                    col += 1
                    If Not IsNothing(prg.elev_bot_actual) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.elev_bot_actual, Double)
                    col += 1
                    If Not IsNothing(prg.elev_bot_eff) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.elev_bot_eff, Double)
                    col += 1
                    If Not IsNothing(prg.elev_top_actual) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.elev_top_actual, Double)
                    col += 1
                    If Not IsNothing(prg.elev_top_eff) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.elev_top_eff, Double)
                    col += 1
                    If Not IsNothing(prg.reinf_db_id) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.reinf_db_id, Integer)
                    col += 1
                    If Not IsNothing(prg.qty) Then .Worksheets("Reinf Groups (SAPI)").Cells(row, col).Value = CType(prg.qty, Integer)


                    'Individual Reinforcement Details 
                    drow = 3
                    For Each prd As PoleReinfDetail In prg.reinf_ids

                        'If prd.reinf_group_id = prg.reinf_group_id Then '--Need help setting up subtable. reinf_group_id is not a field within PoleReinfDetail

                        dcol = 1

                            If Not IsNothing(prg.reinf_group_id) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prg.reinf_group_id, Integer)
                            dcol += 1
                            If Not IsNothing(prd.reinf_id) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.reinf_id, Integer)
                            dcol += 1
                            If Not IsNothing(prd.pole_flat) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.pole_flat, Integer)
                            dcol += 1
                            If Not IsNothing(prd.horizontal_offset) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.horizontal_offset, Double)
                            dcol += 1
                            If Not IsNothing(prd.rotation) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = CType(prd.rotation, Double)
                            dcol += 1
                            If Not IsNothing(prd.note) Then .Worksheets("Reinf ID (SAPI)").Cells(drow, dcol).Value = prd.note

                            drow += 1

                        'End If

                    Next

                    row += 1

                Next

                row = 3

                'Interference Groups
                For Each pig As PoleIntGroup In cp.int_groups

                    col = 1

                    If Not IsNothing(cp.pole_structure_id) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(cp.pole_structure_id, Integer)
                    col += 1
                    If Not IsNothing(pig.interference_group_id) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.interference_group_id, Integer)
                    col += 1
                    If Not IsNothing(pig.elev_bot) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.elev_bot, Double)
                    col += 1
                    If Not IsNothing(pig.elev_top) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.elev_top, Double)
                    col += 1
                    If Not IsNothing(pig.width) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = CType(pig.width, Double)
                    col += 1
                    If Not IsNothing(pig.description) Then .Worksheets("Interference Groups (SAPI)").Cells(row, col).Value = pig.description

                    'Individual Interference Details
                    drow = 3
                    For Each pid As PoleIntDetail In pig.int_ids

                        'If pid.interference_id = pig.interference_group_id Then '--Need help setting up subtable. reinf_group_id is not a field within PoleReinfDetail

                        dcol = 1

                            If Not IsNothing(pig.interference_group_id) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pig.interference_group_id, Integer)
                            dcol += 1
                            If Not IsNothing(pid.interference_id) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pid.interference_id, Integer)
                            dcol += 1
                            If Not IsNothing(pid.pole_flat) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pid.pole_flat, Integer)
                            dcol += 1
                            If Not IsNothing(pid.horizontal_offset) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pid.horizontal_offset, Double)
                            dcol += 1
                            If Not IsNothing(pid.rotation) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = CType(pid.rotation, Double)
                            dcol += 1
                            If Not IsNothing(pid.note) Then .Worksheets("Interference ID (SAPI)").Cells(drow, dcol).Value = pid.note

                            drow += 1

                        'End If

                    Next

                    row += 1

                Next

                row = 3

                'Reinforced Pole Section Results
                For Each prr As PoleReinfResults In cp.reinf_section_results

                    col = 1

                    If Not IsNothing(prr.work_order_seq_num) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.work_order_seq_num, Double)
                    col += 1
                    If Not IsNothing(prr.section_id) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.section_id, Integer)
                    col += 1
                    If Not IsNothing(prr.analysis_section_id) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.analysis_section_id, Integer)
                    col += 1
                    If Not IsNothing(prr.reinf_group_id) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.reinf_group_id, Integer)
                    col += 1
                    If Not IsNothing(prr.result_lkup_value) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.result_lkup_value, Integer)
                    col += 1
                    If Not IsNothing(prr.rating) Then .Worksheets("Reinf Results (SAPI)").Cells(row, col).Value = CType(prr.rating, Double)

                    row += 1

                Next

                row = 3

                'Custom Reinforcement Properties
                For Each pr As PropReinf In cp.reinfs

                    col = 1

                    If Not IsNothing(cp.pole_structure_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(cp.pole_structure_id, Integer)
                    col += 1
                    If Not IsNothing(pr.reinf_db_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.reinf_db_id, Integer)
                    col += 1
                    If Not IsNothing(pr.reinf_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.reinf_id, Integer)
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
                    If Not IsNothing(pr.istension) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.istension
                    col += 1
                    If Not IsNothing(pr.matl_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.matl_id, Integer)
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
                    If Not IsNothing(pr.bolt_id_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_id_bot, Integer)
                    col += 1
                    If Not IsNothing(pr.bolt_N_or_X_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.bolt_N_or_X_bot
                    col += 1
                    If Not IsNothing(pr.bolt_num_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_num_bot, Integer)
                    col += 1
                    If Not IsNothing(pr.bolt_spacing_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_spacing_bot, Double)
                    col += 1
                    If Not IsNothing(pr.bolt_edge_dist_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_edge_dist_bot, Double)
                    col += 1
                    If Not IsNothing(pr.FlangeOrBP_connected_bot) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.FlangeOrBP_connected_bot
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
                    If Not IsNothing(pr.top_bot_connections_symmetrical) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.top_bot_connections_symmetrical
                    col += 1
                    If Not IsNothing(pr.connection_type_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.connection_type_top
                    col += 1
                    If Not IsNothing(pr.connection_cap_revF_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.connection_cap_revF_top, Double)
                    col += 1
                    If Not IsNothing(pr.connection_cap_revG_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.connection_cap_revG_top, Double)
                    col += 1
                    If Not IsNothing(pr.connection_cap_revH_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.connection_cap_revH_top, Double)
                    col += 1
                    If Not IsNothing(pr.bolt_id_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_id_top, Integer)
                    col += 1
                    If Not IsNothing(pr.bolt_N_or_X_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.bolt_N_or_X_top
                    col += 1
                    If Not IsNothing(pr.bolt_num_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_num_top, Integer)
                    col += 1
                    If Not IsNothing(pr.bolt_spacing_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_spacing_top, Double)
                    col += 1
                    If Not IsNothing(pr.bolt_edge_dist_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pr.bolt_edge_dist_top, Double)
                    col += 1
                    If Not IsNothing(pr.FlangeOrBP_connected_top) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pr.FlangeOrBP_connected_top
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

                    row += 1

                Next

                row = 3

                'Custom Bolt Properties
                For Each pb As PropBolt In cp.bolts

                    col = 1

                    If Not IsNothing(cp.pole_structure_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(cp.pole_structure_id, Integer)
                    col += 1
                    If Not IsNothing(pb.bolt_db_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.bolt_db_id, Integer)
                    col += 1
                    If Not IsNothing(pb.bolt_id) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.bolt_id, Integer)
                    col += 1
                    If Not IsNothing(pb.name) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pb.name
                    col += 1
                    If Not IsNothing(pb.description) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pb.description
                    col += 1
                    If Not IsNothing(pb.diam) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.diam, Double)
                    col += 1
                    If Not IsNothing(pb.area) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.area, Double)
                    col += 1
                    If Not IsNothing(pb.fu_bolt) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.fu_bolt, Double)
                    col += 1
                    If Not IsNothing(pb.sleeve_diam_out) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.sleeve_diam_out, Double)
                    col += 1
                    If Not IsNothing(pb.sleeve_diam_in) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.sleeve_diam_in, Double)
                    col += 1
                    If Not IsNothing(pb.fu_sleeve) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.fu_sleeve, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_n_sleeve_shear_revF) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.bolt_n_sleeve_shear_revF, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_x_sleeve_shear_revF) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.bolt_x_sleeve_shear_revF, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_n_sleeve_shear_revG) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.bolt_n_sleeve_shear_revG, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_x_sleeve_shear_revG) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.bolt_x_sleeve_shear_revG, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_n_sleeve_shear_revH) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.bolt_n_sleeve_shear_revH, Double)
                    col += 1
                    If Not IsNothing(pb.bolt_x_sleeve_shear_revH) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = CType(pb.bolt_x_sleeve_shear_revH, Double)
                    col += 1
                    If Not IsNothing(pb.rb_applied_revH) Then .Worksheets("Reinforcements (SAPI)").Cells(row, col).Value = pb.rb_applied_revH

                    row += 1

                Next

                row = 3

                'Custom Material Properties
                For Each pm As PropMatl In cp.matls

                    col = 1

                    If Not IsNothing(cp.pole_structure_id) Then .Worksheets("Materials (SAPI)").Range("").Value = CType(cp.pole_structure_id, Integer)
                    col += 1
                    If Not IsNothing(pm.matl_db_id) Then .Worksheets("Materials (SAPI)").Range("").Value = CType(pm.matl_db_id, Integer)
                    col += 1
                    If Not IsNothing(pm.matl_id) Then .Worksheets("Materials (SAPI)").Range("").Value = CType(pm.matl_id, Integer)
                    col += 1
                    If Not IsNothing(pm.name) Then .Worksheets("Materials (SAPI)").Range("").Value = pm.name
                    col += 1
                    If Not IsNothing(pm.fy) Then .Worksheets("Materials (SAPI)").Range("").Value = CType(pm.fy, Double)
                    col += 1
                    If Not IsNothing(pm.fu) Then .Worksheets("Materials (SAPI)").Range("").Value = CType(pm.fu, Double)

                    row += 1

                Next

            End With

            SaveAndCloseCCIpole()

        Next

        ''Worksheet Change Events
        'If pf.pile_group_config = "Circular" Then
        '    .Worksheets("Moment of Inertia").Visible = False
        '    .Worksheets("Moment of Inertia (Circle)").Visible = True
        'Else
        '    .Worksheets("Moment of Inertia").Visible = True
        '    .Worksheets("Moment of Inertia (Circle)").Visible = False
        'End If

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

    Private Sub LoadNewPole()
        NewCCIpoleWb.LoadDocument(CCIpoleTemplatePath, CCIpoleFileType)
        NewCCIpoleWb.BeginUpdate()
    End Sub

    Private Sub SaveAndCloseCCIpole()
        NewCCIpoleWb.Calculate()
        NewCCIpoleWb.EndUpdate()
        NewCCIpoleWb.SaveDocument(ExcelFilePath, CCIpoleFileType)
    End Sub
#End Region

#Region "SQL Insert Statements"
    Private Function InsertPoleCriteria(ByVal pc As PoleCriteria) As String
        Dim insertString As String = ""

        insertString += "@PoleCriteriaID"
        insertString += "," & IIf(IsNothing(pc.criteria_id), "Null", pc.criteria_id.ToString)
        insertString += "," & IIf(IsNothing(pc.upper_structure_type), "Null", "'" & pc.upper_structure_type.ToString & "'")
        insertString += "," & IIf(IsNothing(pc.analysis_deg), "Null", pc.analysis_deg.ToString)
        insertString += "," & IIf(IsNothing(pc.geom_increment_length), "Null", pc.geom_increment_length.ToString)
        insertString += "," & IIf(IsNothing(pc.vnum), "Null", "'" & pc.vnum.ToString & "'")
        insertString += "," & IIf(IsNothing(pc.check_connections), "Null", "'" & pc.check_connections.ToString & "'")
        insertString += "," & IIf(IsNothing(pc.hole_deformation), "Null", "'" & pc.hole_deformation.ToString & "'")
        insertString += "," & IIf(IsNothing(pc.ineff_mod_check), "Null", "'" & pc.ineff_mod_check.ToString & "'")
        insertString += "," & IIf(IsNothing(pc.modified), "Null", "'" & pc.modified.ToString & "'")

        Return insertString
    End Function

    Private Function InsertPoleSection(ByVal ps As PoleSection) As String
        Dim insertString As String = ""

        insertString += "@PoleSectionID"
        'insertString += "," & IIf(IsNothing(ps.section_id), "Null", ps.section_id.ToString)
        insertString += "," & IIf(IsNothing(ps.local_section_id), "Null", ps.local_section_id.ToString)
        insertString += "," & IIf(IsNothing(ps.elev_bot), "Null", ps.elev_bot.ToString)
        insertString += "," & IIf(IsNothing(ps.elev_top), "Null", ps.elev_top.ToString)
        insertString += "," & IIf(IsNothing(ps.length_section), "Null", ps.length_section.ToString)
        insertString += "," & IIf(IsNothing(ps.length_splice), "Null", ps.length_splice.ToString)
        insertString += "," & IIf(IsNothing(ps.num_sides), "Null", ps.num_sides.ToString)
        insertString += "," & IIf(IsNothing(ps.diam_bot), "Null", ps.diam_bot.ToString)
        insertString += "," & IIf(IsNothing(ps.diam_top), "Null", ps.diam_top.ToString)
        insertString += "," & IIf(IsNothing(ps.wall_thickness), "Null", ps.wall_thickness.ToString)
        insertString += "," & IIf(IsNothing(ps.bend_radius), "Null", ps.bend_radius.ToString)
        insertString += "," & "@MatlID" 'IIf(IsNothing(ps.steel_grade_id), "Null", ps.steel_grade_id.ToString)
        insertString += "," & IIf(IsNothing(ps.pole_type), "Null", "'" & ps.pole_type.ToString & "'")
        insertString += "," & IIf(IsNothing(ps.section_name), "Null", "'" & ps.section_name.ToString & "'")
        insertString += "," & IIf(IsNothing(ps.socket_length), "Null", ps.socket_length.ToString)
        insertString += "," & IIf(IsNothing(ps.weight_mult), "Null", ps.weight_mult.ToString)
        insertString += "," & IIf(IsNothing(ps.wp_mult), "Null", ps.wp_mult.ToString)
        insertString += "," & IIf(IsNothing(ps.af_factor), "Null", ps.af_factor.ToString)
        insertString += "," & IIf(IsNothing(ps.ar_factor), "Null", ps.ar_factor.ToString)
        insertString += "," & IIf(IsNothing(ps.round_area_ratio), "Null", ps.round_area_ratio.ToString)
        insertString += "," & IIf(IsNothing(ps.flat_area_ratio), "Null", ps.flat_area_ratio.ToString)

        Return insertString
    End Function

    Private Function InsertPoleReinfSection(ByVal prs As PoleReinfSection) As String
        Dim insertString As String = ""

        insertString += "@PoleReinfSectionID"
        insertString += "," & IIf(IsNothing(prs.section_ID), "Null", prs.section_ID.ToString)
        insertString += "," & IIf(IsNothing(prs.local_section_id), "Null", prs.local_section_id.ToString)
        insertString += "," & IIf(IsNothing(prs.elev_bot), "Null", prs.elev_bot.ToString)
        insertString += "," & IIf(IsNothing(prs.elev_top), "Null", prs.elev_top.ToString)
        insertString += "," & IIf(IsNothing(prs.length_section), "Null", prs.length_section.ToString)
        insertString += "," & IIf(IsNothing(prs.length_splice), "Null", prs.length_splice.ToString)
        insertString += "," & IIf(IsNothing(prs.num_sides), "Null", prs.num_sides.ToString)
        insertString += "," & IIf(IsNothing(prs.diam_bot), "Null", prs.diam_bot.ToString)
        insertString += "," & IIf(IsNothing(prs.diam_top), "Null", prs.diam_top.ToString)
        insertString += "," & IIf(IsNothing(prs.wall_thickness), "Null", prs.wall_thickness.ToString)
        insertString += "," & IIf(IsNothing(prs.bend_radius), "Null", prs.bend_radius.ToString)
        insertString += "," & IIf(IsNothing(prs.steel_grade_id), "Null", prs.steel_grade_id.ToString)
        insertString += "," & IIf(IsNothing(prs.pole_type), "Null", "'" & prs.pole_type.ToString & "'")
        insertString += "," & IIf(IsNothing(prs.weight_mult), "Null", prs.weight_mult.ToString)
        insertString += "," & IIf(IsNothing(prs.section_name), "Null", "'" & prs.section_name.ToString & "'")
        insertString += "," & IIf(IsNothing(prs.socket_length), "Null", prs.socket_length.ToString)
        insertString += "," & IIf(IsNothing(prs.wp_mult), "Null", prs.wp_mult.ToString)
        insertString += "," & IIf(IsNothing(prs.af_factor), "Null", prs.af_factor.ToString)
        insertString += "," & IIf(IsNothing(prs.ar_factor), "Null", prs.ar_factor.ToString)
        insertString += "," & IIf(IsNothing(prs.round_area_ratio), "Null", prs.round_area_ratio.ToString)
        insertString += "," & IIf(IsNothing(prs.flat_area_ratio), "Null", prs.flat_area_ratio.ToString)

        Return insertString
    End Function

    Private Function InsertPoleReinfGroup(ByVal prg As PoleReinfGroup) As String
        Dim insertString As String = ""

        insertString += "@PoleReinfGroupID"
        insertString += "," & IIf(IsNothing(prg.reinf_group_id), "Null", prg.reinf_group_id.ToString)
        insertString += "," & IIf(IsNothing(prg.elev_bot_actual), "Null", prg.elev_bot_actual.ToString)
        insertString += "," & IIf(IsNothing(prg.elev_bot_eff), "Null", prg.elev_bot_eff.ToString)
        insertString += "," & IIf(IsNothing(prg.elev_top_actual), "Null", prg.elev_top_actual.ToString)
        insertString += "," & IIf(IsNothing(prg.elev_top_eff), "Null", prg.elev_top_eff.ToString)
        insertString += "," & "@TypeID" 'IIf(IsNothing(prg.reinf_db_id), "Null", prg.reinf_db_id.ToString)  IF 18 or Less then set to NULL in sub sub query


        Return insertString
    End Function

    Private Function InsertPoleReinfDetail(ByVal prd As PoleReinfDetail) As String
        Dim insertString As String = ""

        insertString += "@PoleReinfDetailID"
        insertString += "," & IIf(IsNothing(prd.reinf_id), "Null", prd.reinf_id.ToString)
        insertString += "," & IIf(IsNothing(prd.pole_flat), "Null", prd.pole_flat.ToString)
        insertString += "," & IIf(IsNothing(prd.horizontal_offset), "Null", prd.horizontal_offset.ToString)
        insertString += "," & IIf(IsNothing(prd.rotation), "Null", prd.rotation.ToString)
        insertString += "," & IIf(IsNothing(prd.note), "Null", "'" & prd.note.ToString & "'")

        Return insertString
    End Function

    Private Function InsertPoleIntGroup(ByVal pig As PoleIntGroup) As String
        Dim insertString As String = ""

        insertString += "@PoleIntGroupID"
        insertString += "," & IIf(IsNothing(pig.interference_group_id), "Null", pig.interference_group_id.ToString)
        insertString += "," & IIf(IsNothing(pig.elev_bot), "Null", pig.elev_bot.ToString)
        insertString += "," & IIf(IsNothing(pig.elev_top), "Null", pig.elev_top.ToString)
        insertString += "," & IIf(IsNothing(pig.width), "Null", pig.width.ToString)
        insertString += "," & IIf(IsNothing(pig.description), "Null", "'" & pig.description.ToString & "'")

        Return insertString
    End Function

    Private Function InsertPoleIntDetail(ByVal pid As PoleIntDetail) As String
        Dim insertString As String = ""

        insertString += "@IntDetailID"
        insertString += "," & IIf(IsNothing(pid.interference_id), "Null", pid.interference_id.ToString)
        insertString += "," & IIf(IsNothing(pid.pole_flat), "Null", pid.pole_flat.ToString)
        insertString += "," & IIf(IsNothing(pid.horizontal_offset), "Null", pid.horizontal_offset.ToString)
        insertString += "," & IIf(IsNothing(pid.rotation), "Null", pid.rotation.ToString)
        insertString += "," & IIf(IsNothing(pid.note), "Null", "'" & pid.note.ToString & "'")

        Return insertString
    End Function

    Private Function InsertPoleReinfResults(ByVal prr As PoleReinfResults) As String
        Dim insertString As String = ""

        insertString += "@PoleReinfResultID"
        insertString += "," & IIf(IsNothing(prr.section_id), "Null", prr.section_id.ToString)
        insertString += "," & IIf(IsNothing(prr.work_order_seq_num), "Null", prr.work_order_seq_num.ToString)
        insertString += "," & IIf(IsNothing(prr.reinf_group_id), "Null", prr.reinf_group_id.ToString)
        insertString += "," & IIf(IsNothing(prr.result_lkup_value), "Null", prr.result_lkup_value.ToString)
        insertString += "," & IIf(IsNothing(prr.rating), "Null", prr.rating.ToString)

        Return insertString
    End Function

    Private Function InsertPropReinf(ByVal pr As PropReinf) As String
        Dim insertString As String = ""

        insertString += "@ReinfID"
        insertString += "," & IIf(IsNothing(pr.reinf_db_id), "Null", pr.reinf_db_id.ToString)
        insertString += "," & IIf(IsNothing(pr.name), "Null", "'" & pr.name.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.type), "Null", "'" & pr.type.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.b), "Null", pr.b.ToString)
        insertString += "," & IIf(IsNothing(pr.h), "Null", pr.h.ToString)
        insertString += "," & IIf(IsNothing(pr.sr_diam), "Null", pr.sr_diam.ToString)
        insertString += "," & IIf(IsNothing(pr.channel_thkns_web), "Null", pr.channel_thkns_web.ToString)
        insertString += "," & IIf(IsNothing(pr.channel_thkns_flange), "Null", pr.channel_thkns_flange.ToString)
        insertString += "," & IIf(IsNothing(pr.channel_eo), "Null", pr.channel_eo.ToString)
        insertString += "," & IIf(IsNothing(pr.channel_J), "Null", pr.channel_J.ToString)
        insertString += "," & IIf(IsNothing(pr.channel_Cw), "Null", pr.channel_Cw.ToString)
        insertString += "," & IIf(IsNothing(pr.area_gross), "Null", pr.area_gross.ToString)
        insertString += "," & IIf(IsNothing(pr.centroid), "Null", pr.centroid.ToString)
        insertString += "," & IIf(IsNothing(pr.istension), "Null", "'" & pr.istension.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.matl_id), "Null", pr.matl_id.ToString)
        insertString += "," & IIf(IsNothing(pr.Ix), "Null", pr.Ix.ToString)
        insertString += "," & IIf(IsNothing(pr.Iy), "Null", pr.Iy.ToString)
        insertString += "," & IIf(IsNothing(pr.Lu), "Null", pr.Lu.ToString)
        insertString += "," & IIf(IsNothing(pr.Kx), "Null", pr.Kx.ToString)
        insertString += "," & IIf(IsNothing(pr.Ky), "Null", pr.Ky.ToString)
        insertString += "," & IIf(IsNothing(pr.bolt_hole_size), "Null", pr.bolt_hole_size.ToString)
        insertString += "," & IIf(IsNothing(pr.area_net), "Null", pr.area_net.ToString)
        insertString += "," & IIf(IsNothing(pr.shear_lag), "Null", pr.shear_lag.ToString)
        insertString += "," & IIf(IsNothing(pr.connection_type_bot), "Null", "'" & pr.connection_type_bot.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.connection_cap_revF_bot), "Null", pr.connection_cap_revF_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.connection_cap_revG_bot), "Null", pr.connection_cap_revG_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.connection_cap_revH_bot), "Null", pr.connection_cap_revH_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.bolt_id_bot), "Null", pr.bolt_id_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.bolt_N_or_X_bot), "Null", "'" & pr.bolt_N_or_X_bot.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.bolt_num_bot), "Null", pr.bolt_num_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.bolt_spacing_bot), "Null", pr.bolt_spacing_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.bolt_edge_dist_bot), "Null", pr.bolt_edge_dist_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.FlangeOrBP_connected_bot), "Null", "'" & pr.FlangeOrBP_connected_bot.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.weld_grade_bot), "Null", pr.weld_grade_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_trans_type_bot), "Null", "'" & pr.weld_trans_type_bot.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.weld_trans_length_bot), "Null", pr.weld_trans_length_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_groove_depth_bot), "Null", pr.weld_groove_depth_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_groove_angle_bot), "Null", pr.weld_groove_angle_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_trans_fillet_size_bot), "Null", pr.weld_trans_fillet_size_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_trans_eff_throat_bot), "Null", pr.weld_trans_eff_throat_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_long_type_bot), "Null", "'" & pr.weld_long_type_bot.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.weld_long_length_bot), "Null", pr.weld_long_length_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_long_fillet_size_bot), "Null", pr.weld_long_fillet_size_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_long_eff_throat_bot), "Null", pr.weld_long_eff_throat_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.top_bot_connections_symmetrical), "Null", "'" & pr.top_bot_connections_symmetrical.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.connection_type_top), "Null", "'" & pr.connection_type_top.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.connection_cap_revF_top), "Null", pr.connection_cap_revF_top.ToString)
        insertString += "," & IIf(IsNothing(pr.connection_cap_revG_top), "Null", pr.connection_cap_revG_top.ToString)
        insertString += "," & IIf(IsNothing(pr.connection_cap_revH_top), "Null", pr.connection_cap_revH_top.ToString)
        insertString += "," & IIf(IsNothing(pr.bolt_id_top), "Null", pr.bolt_id_top.ToString)
        insertString += "," & IIf(IsNothing(pr.bolt_N_or_X_top), "Null", "'" & pr.bolt_N_or_X_top.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.bolt_num_top), "Null", pr.bolt_num_top.ToString)
        insertString += "," & IIf(IsNothing(pr.bolt_spacing_top), "Null", pr.bolt_spacing_top.ToString)
        insertString += "," & IIf(IsNothing(pr.bolt_edge_dist_top), "Null", pr.bolt_edge_dist_top.ToString)
        insertString += "," & IIf(IsNothing(pr.FlangeOrBP_connected_top), "Null", "'" & pr.FlangeOrBP_connected_top.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.weld_grade_top), "Null", pr.weld_grade_top.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_trans_type_top), "Null", "'" & pr.weld_trans_type_top.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.weld_trans_length_top), "Null", pr.weld_trans_length_top.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_groove_depth_top), "Null", pr.weld_groove_depth_top.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_groove_angle_top), "Null", pr.weld_groove_angle_top.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_trans_fillet_size_top), "Null", pr.weld_trans_fillet_size_top.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_trans_eff_throat_top), "Null", pr.weld_trans_eff_throat_top.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_long_type_top), "Null", "'" & pr.weld_long_type_top.ToString & "'")
        insertString += "," & IIf(IsNothing(pr.weld_long_length_top), "Null", pr.weld_long_length_top.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_long_fillet_size_top), "Null", pr.weld_long_fillet_size_top.ToString)
        insertString += "," & IIf(IsNothing(pr.weld_long_eff_throat_top), "Null", pr.weld_long_eff_throat_top.ToString)
        insertString += "," & IIf(IsNothing(pr.conn_length_bot), "Null", pr.conn_length_bot.ToString)
        insertString += "," & IIf(IsNothing(pr.conn_length_top), "Null", pr.conn_length_top.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_comp_xx_f), "Null", pr.cap_comp_xx_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_comp_yy_f), "Null", pr.cap_comp_yy_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_tens_yield_f), "Null", pr.cap_tens_yield_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_tens_rupture_f), "Null", pr.cap_tens_rupture_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_shear_f), "Null", pr.cap_shear_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_bot_f), "Null", pr.cap_bolt_shear_bot_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_top_f), "Null", pr.cap_bolt_shear_top_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_f), "Null", pr.cap_boltshaft_bearing_nodeform_bot_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_f), "Null", pr.cap_boltshaft_bearing_deform_bot_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_f), "Null", pr.cap_boltshaft_bearing_nodeform_top_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_f), "Null", pr.cap_boltshaft_bearing_deform_top_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_f), "Null", pr.cap_boltreinf_bearing_nodeform_bot_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_f), "Null", pr.cap_boltreinf_bearing_deform_bot_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_f), "Null", pr.cap_boltreinf_bearing_nodeform_top_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_f), "Null", pr.cap_boltreinf_bearing_deform_top_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_bot_f), "Null", pr.cap_weld_trans_bot_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_long_bot_f), "Null", pr.cap_weld_long_bot_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_top_f), "Null", pr.cap_weld_trans_top_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_long_top_f), "Null", pr.cap_weld_long_top_f.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_comp_xx_g), "Null", pr.cap_comp_xx_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_comp_yy_g), "Null", pr.cap_comp_yy_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_tens_yield_g), "Null", pr.cap_tens_yield_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_tens_rupture_g), "Null", pr.cap_tens_rupture_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_shear_g), "Null", pr.cap_shear_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_bot_g), "Null", pr.cap_bolt_shear_bot_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_top_g), "Null", pr.cap_bolt_shear_top_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_g), "Null", pr.cap_boltshaft_bearing_nodeform_bot_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_g), "Null", pr.cap_boltshaft_bearing_deform_bot_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_g), "Null", pr.cap_boltshaft_bearing_nodeform_top_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_g), "Null", pr.cap_boltshaft_bearing_deform_top_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_g), "Null", pr.cap_boltreinf_bearing_nodeform_bot_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_g), "Null", pr.cap_boltreinf_bearing_deform_bot_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_g), "Null", pr.cap_boltreinf_bearing_nodeform_top_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_g), "Null", pr.cap_boltreinf_bearing_deform_top_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_bot_g), "Null", pr.cap_weld_trans_bot_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_long_bot_g), "Null", pr.cap_weld_long_bot_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_top_g), "Null", pr.cap_weld_trans_top_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_long_top_g), "Null", pr.cap_weld_long_top_g.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_comp_xx_h), "Null", pr.cap_comp_xx_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_comp_yy_h), "Null", pr.cap_comp_yy_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_tens_yield_h), "Null", pr.cap_tens_yield_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_tens_rupture_h), "Null", pr.cap_tens_rupture_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_shear_h), "Null", pr.cap_shear_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_bot_h), "Null", pr.cap_bolt_shear_bot_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_bolt_shear_top_h), "Null", pr.cap_bolt_shear_top_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_h), "Null", pr.cap_boltshaft_bearing_nodeform_bot_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_h), "Null", pr.cap_boltshaft_bearing_deform_bot_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_h), "Null", pr.cap_boltshaft_bearing_nodeform_top_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_h), "Null", pr.cap_boltshaft_bearing_deform_top_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_h), "Null", pr.cap_boltreinf_bearing_nodeform_bot_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_h), "Null", pr.cap_boltreinf_bearing_deform_bot_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_h), "Null", pr.cap_boltreinf_bearing_nodeform_top_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_h), "Null", pr.cap_boltreinf_bearing_deform_top_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_bot_h), "Null", pr.cap_weld_trans_bot_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_long_bot_h), "Null", pr.cap_weld_long_bot_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_trans_top_h), "Null", pr.cap_weld_trans_top_h.ToString)
        insertString += "," & IIf(IsNothing(pr.cap_weld_long_top_h), "Null", pr.cap_weld_long_top_h.ToString)

        Return insertString
    End Function

    Private Function InsertPropBolt(ByVal pb As PropBolt) As String
        Dim insertString As String = ""

        insertString += "@BoltID"
        insertString += "," & IIf(IsNothing(pb.bolt_db_id), "Null", pb.bolt_db_id.ToString)
        insertString += "," & IIf(IsNothing(pb.name), "Null", "'" & pb.name.ToString & "'")
        insertString += "," & IIf(IsNothing(pb.description), "Null", "'" & pb.description.ToString & "'")
        insertString += "," & IIf(IsNothing(pb.diam), "Null", pb.diam.ToString)
        insertString += "," & IIf(IsNothing(pb.area), "Null", pb.area.ToString)
        insertString += "," & IIf(IsNothing(pb.fu_bolt), "Null", pb.fu_bolt.ToString)
        insertString += "," & IIf(IsNothing(pb.sleeve_diam_out), "Null", pb.sleeve_diam_out.ToString)
        insertString += "," & IIf(IsNothing(pb.sleeve_diam_in), "Null", pb.sleeve_diam_in.ToString)
        insertString += "," & IIf(IsNothing(pb.fu_sleeve), "Null", pb.fu_sleeve.ToString)
        insertString += "," & IIf(IsNothing(pb.bolt_n_sleeve_shear_revF), "Null", pb.bolt_n_sleeve_shear_revF.ToString)
        insertString += "," & IIf(IsNothing(pb.bolt_x_sleeve_shear_revF), "Null", pb.bolt_x_sleeve_shear_revF.ToString)
        insertString += "," & IIf(IsNothing(pb.bolt_n_sleeve_shear_revG), "Null", pb.bolt_n_sleeve_shear_revG.ToString)
        insertString += "," & IIf(IsNothing(pb.bolt_x_sleeve_shear_revG), "Null", pb.bolt_x_sleeve_shear_revG.ToString)
        insertString += "," & IIf(IsNothing(pb.bolt_n_sleeve_shear_revH), "Null", pb.bolt_n_sleeve_shear_revH.ToString)
        insertString += "," & IIf(IsNothing(pb.bolt_x_sleeve_shear_revH), "Null", pb.bolt_x_sleeve_shear_revH.ToString)
        insertString += "," & IIf(IsNothing(pb.rb_applied_revH), "Null", "'" & pb.rb_applied_revH.ToString & "'")

        Return insertString
    End Function

    Private Function InsertPropMatl(ByVal pm As PropMatl) As String
        Dim insertString As String = ""

        insertString += "@MatlID"
        insertString += "," & IIf(IsNothing(pm.matl_db_id), "Null", pm.matl_db_id.ToString)
        insertString += "," & IIf(IsNothing(pm.name), "Null", "'" & pm.name.ToString & "'")
        insertString += "," & IIf(IsNothing(pm.fy), "Null", pm.fy.ToString)
        insertString += "," & IIf(IsNothing(pm.fu), "Null", pm.fu.ToString)

        Return insertString
    End Function

#End Region

#Region "SQL Update Statements"
    Private Function UpdatePoleCriteria(ByVal pc As PoleCriteria) As String
        Dim updateString As String = ""

        updateString += "UPDATE pole_analysis_criteria SET "
        updateString += " criteria_id=" & IIf(IsNothing(pc.criteria_id), "Null", pc.criteria_id.ToString)
        updateString += ", upper_structure_type=" & IIf(IsNothing(pc.upper_structure_type), "Null", "'" & pc.upper_structure_type.ToString & "'")
        updateString += ", analysis_deg=" & IIf(IsNothing(pc.analysis_deg), "Null", pc.analysis_deg.ToString)
        updateString += ", geom_increment_length=" & IIf(IsNothing(pc.geom_increment_length), "Null", pc.geom_increment_length.ToString)
        updateString += ", vnum=" & IIf(IsNothing(pc.vnum), "Null", "'" & pc.vnum.ToString & "'")
        updateString += ", check_connections=" & IIf(IsNothing(pc.check_connections), "Null", "'" & pc.check_connections.ToString & "'")
        updateString += ", hole_deformation=" & IIf(IsNothing(pc.hole_deformation), "Null", "'" & pc.hole_deformation.ToString & "'")
        updateString += ", ineff_mod_check=" & IIf(IsNothing(pc.ineff_mod_check), "Null", "'" & pc.ineff_mod_check.ToString & "'")
        updateString += ", modified=" & IIf(IsNothing(pc.modified), "Null", "'" & pc.modified.ToString & "'")
        updateString += " WHERE ID = " & pc.criteria_id.ToString

        Return updateString
    End Function

    Private Function UpdatePoleSection(ByVal ps As PoleSection) As String
        Dim updateString As String = ""

        updateString += "UPDATE pole_section SET "
        updateString += " section_id=" & IIf(IsNothing(ps.section_id), "Null", ps.section_id.ToString)
        updateString += ", local_section_id=" & IIf(IsNothing(ps.local_section_id), "Null", ps.local_section_id.ToString)
        updateString += ", elev_bot=" & IIf(IsNothing(ps.elev_bot), "Null", ps.elev_bot.ToString)
        updateString += ", elev_top=" & IIf(IsNothing(ps.elev_top), "Null", ps.elev_top.ToString)
        updateString += ", length_section=" & IIf(IsNothing(ps.length_section), "Null", ps.length_section.ToString)
        updateString += ", length_splice=" & IIf(IsNothing(ps.length_splice), "Null", ps.length_splice.ToString)
        updateString += ", num_sides=" & IIf(IsNothing(ps.num_sides), "Null", ps.num_sides.ToString)
        updateString += ", diam_bot=" & IIf(IsNothing(ps.diam_bot), "Null", ps.diam_bot.ToString)
        updateString += ", diam_top=" & IIf(IsNothing(ps.diam_top), "Null", ps.diam_top.ToString)
        updateString += ", wall_thickness=" & IIf(IsNothing(ps.wall_thickness), "Null", ps.wall_thickness.ToString)
        updateString += ", bend_radius=" & IIf(IsNothing(ps.bend_radius), "Null", ps.bend_radius.ToString)
        updateString += ", steel_grade=" & IIf(IsNothing(ps.steel_grade_id), "Null", ps.steel_grade_id.ToString)
        updateString += ", pole_type=" & IIf(IsNothing(ps.pole_type), "Null", "'" & ps.pole_type.ToString & "'")
        updateString += ", section_name=" & IIf(IsNothing(ps.section_name), "Null", "'" & ps.section_name.ToString & "'")
        updateString += ", socket_length=" & IIf(IsNothing(ps.socket_length), "Null", ps.socket_length.ToString)
        updateString += ", weight_mult=" & IIf(IsNothing(ps.weight_mult), "Null", ps.weight_mult.ToString)
        updateString += ", wp_mult=" & IIf(IsNothing(ps.wp_mult), "Null", ps.wp_mult.ToString)
        updateString += ", af_factor=" & IIf(IsNothing(ps.af_factor), "Null", ps.af_factor.ToString)
        updateString += ", ar_factor=" & IIf(IsNothing(ps.ar_factor), "Null", ps.ar_factor.ToString)
        updateString += ", round_area_ratio=" & IIf(IsNothing(ps.round_area_ratio), "Null", ps.round_area_ratio.ToString)
        updateString += ", flat_area_ratio=" & IIf(IsNothing(ps.flat_area_ratio), "Null", ps.flat_area_ratio.ToString)
        updateString += " WHERE ID = " & ps.section_id.ToString

        Return updateString
    End Function

    Private Function UpdatePoleReinfSection(ByVal prs As PoleReinfSection) As String
        Dim updateString As String = ""

        updateString += "UPDATE pole_reinf_section SET "
        updateString += " section_ID=" & IIf(IsNothing(prs.section_ID), "Null", prs.section_ID.ToString)
        updateString += ", local_section_id=" & IIf(IsNothing(prs.local_section_id), "Null", prs.local_section_id.ToString)
        updateString += ", elev_bot=" & IIf(IsNothing(prs.elev_bot), "Null", prs.elev_bot.ToString)
        updateString += ", elev_top=" & IIf(IsNothing(prs.elev_top), "Null", prs.elev_top.ToString)
        updateString += ", length_section=" & IIf(IsNothing(prs.length_section), "Null", prs.length_section.ToString)
        updateString += ", length_splice=" & IIf(IsNothing(prs.length_splice), "Null", prs.length_splice.ToString)
        updateString += ", num_sides=" & IIf(IsNothing(prs.num_sides), "Null", prs.num_sides.ToString)
        updateString += ", diam_bot=" & IIf(IsNothing(prs.diam_bot), "Null", prs.diam_bot.ToString)
        updateString += ", diam_top=" & IIf(IsNothing(prs.diam_top), "Null", prs.diam_top.ToString)
        updateString += ", wall_thickness=" & IIf(IsNothing(prs.wall_thickness), "Null", prs.wall_thickness.ToString)
        updateString += ", bend_radius=" & IIf(IsNothing(prs.bend_radius), "Null", prs.bend_radius.ToString)
        updateString += ", steel_grade_id=" & IIf(IsNothing(prs.steel_grade_id), "Null", prs.steel_grade_id.ToString)
        updateString += ", pole_type=" & IIf(IsNothing(prs.pole_type), "Null", "'" & prs.pole_type.ToString & "'")
        updateString += ", weight_mult=" & IIf(IsNothing(prs.weight_mult), "Null", prs.weight_mult.ToString)
        updateString += ", section_name=" & IIf(IsNothing(prs.section_name), "Null", "'" & prs.section_name.ToString & "'")
        updateString += ", socket_length=" & IIf(IsNothing(prs.socket_length), "Null", prs.socket_length.ToString)
        updateString += ", wp_mult=" & IIf(IsNothing(prs.wp_mult), "Null", prs.wp_mult.ToString)
        updateString += ", af_factor=" & IIf(IsNothing(prs.af_factor), "Null", prs.af_factor.ToString)
        updateString += ", ar_factor=" & IIf(IsNothing(prs.ar_factor), "Null", prs.ar_factor.ToString)
        updateString += ", round_area_ratio=" & IIf(IsNothing(prs.round_area_ratio), "Null", prs.round_area_ratio.ToString)
        updateString += ", flat_area_ratio=" & IIf(IsNothing(prs.flat_area_ratio), "Null", prs.flat_area_ratio.ToString)
        updateString += " WHERE ID = " & prs.section_ID.ToString

        Return updateString
    End Function

    Private Function UpdatePoleReinfGroup(ByVal prg As PoleReinfGroup) As String
        Dim updateString As String = ""

        updateString += "UPDATE pole_reinf_group SET "
        updateString += " reinf_group_id=" & IIf(IsNothing(prg.reinf_group_id), "Null", prg.reinf_group_id.ToString)
        updateString += ", elev_bot_actual=" & IIf(IsNothing(prg.elev_bot_actual), "Null", prg.elev_bot_actual.ToString)
        updateString += ", elev_bot_eff=" & IIf(IsNothing(prg.elev_bot_eff), "Null", prg.elev_bot_eff.ToString)
        updateString += ", elev_top_actual=" & IIf(IsNothing(prg.elev_top_actual), "Null", prg.elev_top_actual.ToString)
        updateString += ", elev_top_eff=" & IIf(IsNothing(prg.elev_top_eff), "Null", prg.elev_top_eff.ToString)
        updateString += ", reinf_db_id=" & IIf(IsNothing(prg.reinf_db_id), "Null", prg.reinf_db_id.ToString)
        updateString += " WHERE ID = " & prg.reinf_group_id.ToString

        Return updateString
    End Function

    Private Function UpdatePoleReinfDetail(ByVal prd As PoleReinfDetail) As String
        Dim updateString As String = ""

        updateString += "UPDATE pole_reinf_details SET "
        updateString += " reinf_id=" & IIf(IsNothing(prd.reinf_id), "Null", prd.reinf_id.ToString)
        updateString += ", pole_flat=" & IIf(IsNothing(prd.pole_flat), "Null", prd.pole_flat.ToString)
        updateString += ", horizontal_offset=" & IIf(IsNothing(prd.horizontal_offset), "Null", prd.horizontal_offset.ToString)
        updateString += ", rotation=" & IIf(IsNothing(prd.rotation), "Null", prd.rotation.ToString)
        updateString += ", note=" & IIf(IsNothing(prd.note), "Null", "'" & prd.note.ToString & "'")
        updateString += " WHERE ID = " & prd.reinf_id.ToString

        Return updateString
    End Function

    Private Function UpdatePoleIntGroup(ByVal pig As PoleIntGroup) As String
        Dim updateString As String = ""

        updateString += "UPDATE pole_interference_group SET "
        updateString += " interference_group_id=" & IIf(IsNothing(pig.interference_group_id), "Null", pig.interference_group_id.ToString)
        updateString += ", elev_bot=" & IIf(IsNothing(pig.elev_bot), "Null", pig.elev_bot.ToString)
        updateString += ", elev_top=" & IIf(IsNothing(pig.elev_top), "Null", pig.elev_top.ToString)
        updateString += ", width=" & IIf(IsNothing(pig.width), "Null", pig.width.ToString)
        updateString += ", description=" & IIf(IsNothing(pig.description), "Null", "'" & pig.description.ToString & "'")
        updateString += " WHERE ID = " & pig.interference_group_id.ToString

        Return updateString
    End Function

    Private Function UpdatePoleIntDetail(ByVal pid As PoleIntDetail) As String
        Dim updateString As String = ""

        updateString += "UPDATE pole_interference_details SET "
        updateString += " interference_id=" & IIf(IsNothing(pid.interference_id), "Null", pid.interference_id.ToString)
        updateString += ", pole_flat=" & IIf(IsNothing(pid.pole_flat), "Null", pid.pole_flat.ToString)
        updateString += ", horizontal_offset=" & IIf(IsNothing(pid.horizontal_offset), "Null", pid.horizontal_offset.ToString)
        updateString += ", rotation=" & IIf(IsNothing(pid.rotation), "Null", pid.rotation.ToString)
        updateString += ", note=" & IIf(IsNothing(pid.note), "Null", "'" & pid.note.ToString & "'")
        updateString += " WHERE ID = " & pid.interference_id.ToString

        Return updateString
    End Function

    Private Function UpdatePoleReinfResults(ByVal prr As PoleReinfResults) As String
        Dim updateString As String = ""

        updateString += "UPDATE pole_reinf_results SET "
        updateString += " section_id=" & IIf(IsNothing(prr.section_id), "Null", prr.section_id.ToString)
        updateString += ", work_order_seq_num=" & IIf(IsNothing(prr.work_order_seq_num), "Null", prr.work_order_seq_num.ToString)
        updateString += ", reinf_group_id=" & IIf(IsNothing(prr.reinf_group_id), "Null", prr.reinf_group_id.ToString)
        updateString += ", result_lkup_value=" & IIf(IsNothing(prr.result_lkup_value), "Null", prr.result_lkup_value.ToString)
        updateString += ", rating=" & IIf(IsNothing(prr.rating), "Null", prr.rating.ToString)
        updateString += " WHERE ID = " & prr.section_id.ToString

        Return updateString
    End Function

    Private Function UpdatePropReinf(ByVal pr As PropReinf) As String
        Dim updateString As String = ""

        updateString += "UPDATE memb_prop_flat_plate SET "
        updateString += " reinf_db_id=" & IIf(IsNothing(pr.reinf_db_id), "Null", pr.reinf_db_id.ToString)
        updateString += ", name=" & IIf(IsNothing(pr.name), "Null", "'" & pr.name.ToString & "'")
        updateString += ", type=" & IIf(IsNothing(pr.type), "Null", "'" & pr.type.ToString & "'")
        updateString += ", b=" & IIf(IsNothing(pr.b), "Null", pr.b.ToString)
        updateString += ", h=" & IIf(IsNothing(pr.h), "Null", pr.h.ToString)
        updateString += ", sr_diam=" & IIf(IsNothing(pr.sr_diam), "Null", pr.sr_diam.ToString)
        updateString += ", channel_thkns_web=" & IIf(IsNothing(pr.channel_thkns_web), "Null", pr.channel_thkns_web.ToString)
        updateString += ", channel_thkns_flange=" & IIf(IsNothing(pr.channel_thkns_flange), "Null", pr.channel_thkns_flange.ToString)
        updateString += ", channel_eo=" & IIf(IsNothing(pr.channel_eo), "Null", pr.channel_eo.ToString)
        updateString += ", channel_J=" & IIf(IsNothing(pr.channel_J), "Null", pr.channel_J.ToString)
        updateString += ", channel_Cw=" & IIf(IsNothing(pr.channel_Cw), "Null", pr.channel_Cw.ToString)
        updateString += ", area_gross=" & IIf(IsNothing(pr.area_gross), "Null", pr.area_gross.ToString)
        updateString += ", centroid=" & IIf(IsNothing(pr.centroid), "Null", pr.centroid.ToString)
        updateString += ", istension=" & IIf(IsNothing(pr.istension), "Null", "'" & pr.istension.ToString & "'")
        updateString += ", matl_id=" & IIf(IsNothing(pr.matl_id), "Null", pr.matl_id.ToString)
        updateString += ", Ix=" & IIf(IsNothing(pr.Ix), "Null", pr.Ix.ToString)
        updateString += ", Iy=" & IIf(IsNothing(pr.Iy), "Null", pr.Iy.ToString)
        updateString += ", Lu=" & IIf(IsNothing(pr.Lu), "Null", pr.Lu.ToString)
        updateString += ", Kx=" & IIf(IsNothing(pr.Kx), "Null", pr.Kx.ToString)
        updateString += ", Ky=" & IIf(IsNothing(pr.Ky), "Null", pr.Ky.ToString)
        updateString += ", bolt_hole_size=" & IIf(IsNothing(pr.bolt_hole_size), "Null", pr.bolt_hole_size.ToString)
        updateString += ", area_net=" & IIf(IsNothing(pr.area_net), "Null", pr.area_net.ToString)
        updateString += ", shear_lag=" & IIf(IsNothing(pr.shear_lag), "Null", pr.shear_lag.ToString)
        updateString += ", connection_type_bot=" & IIf(IsNothing(pr.connection_type_bot), "Null", "'" & pr.connection_type_bot.ToString & "'")
        updateString += ", connection_cap_revF_bot=" & IIf(IsNothing(pr.connection_cap_revF_bot), "Null", pr.connection_cap_revF_bot.ToString)
        updateString += ", connection_cap_revG_bot=" & IIf(IsNothing(pr.connection_cap_revG_bot), "Null", pr.connection_cap_revG_bot.ToString)
        updateString += ", connection_cap_revH_bot=" & IIf(IsNothing(pr.connection_cap_revH_bot), "Null", pr.connection_cap_revH_bot.ToString)
        updateString += ", bolt_id_bot=" & IIf(IsNothing(pr.bolt_id_bot), "Null", pr.bolt_id_bot.ToString)
        updateString += ", bolt_N_or_X_bot=" & IIf(IsNothing(pr.bolt_N_or_X_bot), "Null", "'" & pr.bolt_N_or_X_bot.ToString & "'")
        updateString += ", bolt_num_bot=" & IIf(IsNothing(pr.bolt_num_bot), "Null", pr.bolt_num_bot.ToString)
        updateString += ", bolt_spacing_bot=" & IIf(IsNothing(pr.bolt_spacing_bot), "Null", pr.bolt_spacing_bot.ToString)
        updateString += ", bolt_edge_dist_bot=" & IIf(IsNothing(pr.bolt_edge_dist_bot), "Null", pr.bolt_edge_dist_bot.ToString)
        updateString += ", FlangeOrBP_connected_bot=" & IIf(IsNothing(pr.FlangeOrBP_connected_bot), "Null", "'" & pr.FlangeOrBP_connected_bot.ToString & "'")
        updateString += ", weld_grade_bot=" & IIf(IsNothing(pr.weld_grade_bot), "Null", pr.weld_grade_bot.ToString)
        updateString += ", weld_trans_type_bot=" & IIf(IsNothing(pr.weld_trans_type_bot), "Null", "'" & pr.weld_trans_type_bot.ToString & "'")
        updateString += ", weld_trans_length_bot=" & IIf(IsNothing(pr.weld_trans_length_bot), "Null", pr.weld_trans_length_bot.ToString)
        updateString += ", weld_groove_depth_bot=" & IIf(IsNothing(pr.weld_groove_depth_bot), "Null", pr.weld_groove_depth_bot.ToString)
        updateString += ", weld_groove_angle_bot=" & IIf(IsNothing(pr.weld_groove_angle_bot), "Null", pr.weld_groove_angle_bot.ToString)
        updateString += ", weld_trans_fillet_size_bot=" & IIf(IsNothing(pr.weld_trans_fillet_size_bot), "Null", pr.weld_trans_fillet_size_bot.ToString)
        updateString += ", weld_trans_eff_throat_bot=" & IIf(IsNothing(pr.weld_trans_eff_throat_bot), "Null", pr.weld_trans_eff_throat_bot.ToString)
        updateString += ", weld_long_type_bot=" & IIf(IsNothing(pr.weld_long_type_bot), "Null", "'" & pr.weld_long_type_bot.ToString & "'")
        updateString += ", weld_long_length_bot=" & IIf(IsNothing(pr.weld_long_length_bot), "Null", pr.weld_long_length_bot.ToString)
        updateString += ", weld_long_fillet_size_bot=" & IIf(IsNothing(pr.weld_long_fillet_size_bot), "Null", pr.weld_long_fillet_size_bot.ToString)
        updateString += ", weld_long_eff_throat_bot=" & IIf(IsNothing(pr.weld_long_eff_throat_bot), "Null", pr.weld_long_eff_throat_bot.ToString)
        updateString += ", top_bot_connections_symmetrical=" & IIf(IsNothing(pr.top_bot_connections_symmetrical), "Null", "'" & pr.top_bot_connections_symmetrical.ToString & "'")
        updateString += ", connection_type_top=" & IIf(IsNothing(pr.connection_type_top), "Null", "'" & pr.connection_type_top.ToString & "'")
        updateString += ", connection_cap_revF_top=" & IIf(IsNothing(pr.connection_cap_revF_top), "Null", pr.connection_cap_revF_top.ToString)
        updateString += ", connection_cap_revG_top=" & IIf(IsNothing(pr.connection_cap_revG_top), "Null", pr.connection_cap_revG_top.ToString)
        updateString += ", connection_cap_revH_top=" & IIf(IsNothing(pr.connection_cap_revH_top), "Null", pr.connection_cap_revH_top.ToString)
        updateString += ", bolt_id_top=" & IIf(IsNothing(pr.bolt_id_top), "Null", pr.bolt_id_top.ToString)
        updateString += ", bolt_N_or_X_top=" & IIf(IsNothing(pr.bolt_N_or_X_top), "Null", "'" & pr.bolt_N_or_X_top.ToString & "'")
        updateString += ", bolt_num_top=" & IIf(IsNothing(pr.bolt_num_top), "Null", pr.bolt_num_top.ToString)
        updateString += ", bolt_spacing_top=" & IIf(IsNothing(pr.bolt_spacing_top), "Null", pr.bolt_spacing_top.ToString)
        updateString += ", bolt_edge_dist_top=" & IIf(IsNothing(pr.bolt_edge_dist_top), "Null", pr.bolt_edge_dist_top.ToString)
        updateString += ", FlangeOrBP_connected_top=" & IIf(IsNothing(pr.FlangeOrBP_connected_top), "Null", "'" & pr.FlangeOrBP_connected_top.ToString & "'")
        updateString += ", weld_grade_top=" & IIf(IsNothing(pr.weld_grade_top), "Null", pr.weld_grade_top.ToString)
        updateString += ", weld_trans_type_top=" & IIf(IsNothing(pr.weld_trans_type_top), "Null", "'" & pr.weld_trans_type_top.ToString & "'")
        updateString += ", weld_trans_length_top=" & IIf(IsNothing(pr.weld_trans_length_top), "Null", pr.weld_trans_length_top.ToString)
        updateString += ", weld_groove_depth_top=" & IIf(IsNothing(pr.weld_groove_depth_top), "Null", pr.weld_groove_depth_top.ToString)
        updateString += ", weld_groove_angle_top=" & IIf(IsNothing(pr.weld_groove_angle_top), "Null", pr.weld_groove_angle_top.ToString)
        updateString += ", weld_trans_fillet_size_top=" & IIf(IsNothing(pr.weld_trans_fillet_size_top), "Null", pr.weld_trans_fillet_size_top.ToString)
        updateString += ", weld_trans_eff_throat_top=" & IIf(IsNothing(pr.weld_trans_eff_throat_top), "Null", pr.weld_trans_eff_throat_top.ToString)
        updateString += ", weld_long_type_top=" & IIf(IsNothing(pr.weld_long_type_top), "Null", "'" & pr.weld_long_type_top.ToString & "'")
        updateString += ", weld_long_length_top=" & IIf(IsNothing(pr.weld_long_length_top), "Null", pr.weld_long_length_top.ToString)
        updateString += ", weld_long_fillet_size_top=" & IIf(IsNothing(pr.weld_long_fillet_size_top), "Null", pr.weld_long_fillet_size_top.ToString)
        updateString += ", weld_long_eff_throat_top=" & IIf(IsNothing(pr.weld_long_eff_throat_top), "Null", pr.weld_long_eff_throat_top.ToString)
        updateString += ", conn_length_bot=" & IIf(IsNothing(pr.conn_length_bot), "Null", pr.conn_length_bot.ToString)
        updateString += ", conn_length_top=" & IIf(IsNothing(pr.conn_length_top), "Null", pr.conn_length_top.ToString)
        updateString += ", cap_comp_xx_f=" & IIf(IsNothing(pr.cap_comp_xx_f), "Null", pr.cap_comp_xx_f.ToString)
        updateString += ", cap_comp_yy_f=" & IIf(IsNothing(pr.cap_comp_yy_f), "Null", pr.cap_comp_yy_f.ToString)
        updateString += ", cap_tens_yield_f=" & IIf(IsNothing(pr.cap_tens_yield_f), "Null", pr.cap_tens_yield_f.ToString)
        updateString += ", cap_tens_rupture_f=" & IIf(IsNothing(pr.cap_tens_rupture_f), "Null", pr.cap_tens_rupture_f.ToString)
        updateString += ", cap_shear_f=" & IIf(IsNothing(pr.cap_shear_f), "Null", pr.cap_shear_f.ToString)
        updateString += ", cap_bolt_shear_bot_f=" & IIf(IsNothing(pr.cap_bolt_shear_bot_f), "Null", pr.cap_bolt_shear_bot_f.ToString)
        updateString += ", cap_bolt_shear_top_f=" & IIf(IsNothing(pr.cap_bolt_shear_top_f), "Null", pr.cap_bolt_shear_top_f.ToString)
        updateString += ", cap_boltshaft_bearing_nodeform_bot_f=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_f), "Null", pr.cap_boltshaft_bearing_nodeform_bot_f.ToString)
        updateString += ", cap_boltshaft_bearing_deform_bot_f=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_f), "Null", pr.cap_boltshaft_bearing_deform_bot_f.ToString)
        updateString += ", cap_boltshaft_bearing_nodeform_top_f=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_f), "Null", pr.cap_boltshaft_bearing_nodeform_top_f.ToString)
        updateString += ", cap_boltshaft_bearing_deform_top_f=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_f), "Null", pr.cap_boltshaft_bearing_deform_top_f.ToString)
        updateString += ", cap_boltreinf_bearing_nodeform_bot_f=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_f), "Null", pr.cap_boltreinf_bearing_nodeform_bot_f.ToString)
        updateString += ", cap_boltreinf_bearing_deform_bot_f=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_f), "Null", pr.cap_boltreinf_bearing_deform_bot_f.ToString)
        updateString += ", cap_boltreinf_bearing_nodeform_top_f=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_f), "Null", pr.cap_boltreinf_bearing_nodeform_top_f.ToString)
        updateString += ", cap_boltreinf_bearing_deform_top_f=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_f), "Null", pr.cap_boltreinf_bearing_deform_top_f.ToString)
        updateString += ", cap_weld_trans_bot_f=" & IIf(IsNothing(pr.cap_weld_trans_bot_f), "Null", pr.cap_weld_trans_bot_f.ToString)
        updateString += ", cap_weld_long_bot_f=" & IIf(IsNothing(pr.cap_weld_long_bot_f), "Null", pr.cap_weld_long_bot_f.ToString)
        updateString += ", cap_weld_trans_top_f=" & IIf(IsNothing(pr.cap_weld_trans_top_f), "Null", pr.cap_weld_trans_top_f.ToString)
        updateString += ", cap_weld_long_top_f=" & IIf(IsNothing(pr.cap_weld_long_top_f), "Null", pr.cap_weld_long_top_f.ToString)
        updateString += ", cap_comp_xx_g=" & IIf(IsNothing(pr.cap_comp_xx_g), "Null", pr.cap_comp_xx_g.ToString)
        updateString += ", cap_comp_yy_g=" & IIf(IsNothing(pr.cap_comp_yy_g), "Null", pr.cap_comp_yy_g.ToString)
        updateString += ", cap_tens_yield_g=" & IIf(IsNothing(pr.cap_tens_yield_g), "Null", pr.cap_tens_yield_g.ToString)
        updateString += ", cap_tens_rupture_g=" & IIf(IsNothing(pr.cap_tens_rupture_g), "Null", pr.cap_tens_rupture_g.ToString)
        updateString += ", cap_shear_g=" & IIf(IsNothing(pr.cap_shear_g), "Null", pr.cap_shear_g.ToString)
        updateString += ", cap_bolt_shear_bot_g=" & IIf(IsNothing(pr.cap_bolt_shear_bot_g), "Null", pr.cap_bolt_shear_bot_g.ToString)
        updateString += ", cap_bolt_shear_top_g=" & IIf(IsNothing(pr.cap_bolt_shear_top_g), "Null", pr.cap_bolt_shear_top_g.ToString)
        updateString += ", cap_boltshaft_bearing_nodeform_bot_g=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_g), "Null", pr.cap_boltshaft_bearing_nodeform_bot_g.ToString)
        updateString += ", cap_boltshaft_bearing_deform_bot_g=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_g), "Null", pr.cap_boltshaft_bearing_deform_bot_g.ToString)
        updateString += ", cap_boltshaft_bearing_nodeform_top_g=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_g), "Null", pr.cap_boltshaft_bearing_nodeform_top_g.ToString)
        updateString += ", cap_boltshaft_bearing_deform_top_g=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_g), "Null", pr.cap_boltshaft_bearing_deform_top_g.ToString)
        updateString += ", cap_boltreinf_bearing_nodeform_bot_g=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_g), "Null", pr.cap_boltreinf_bearing_nodeform_bot_g.ToString)
        updateString += ", cap_boltreinf_bearing_deform_bot_g=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_g), "Null", pr.cap_boltreinf_bearing_deform_bot_g.ToString)
        updateString += ", cap_boltreinf_bearing_nodeform_top_g=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_g), "Null", pr.cap_boltreinf_bearing_nodeform_top_g.ToString)
        updateString += ", cap_boltreinf_bearing_deform_top_g=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_g), "Null", pr.cap_boltreinf_bearing_deform_top_g.ToString)
        updateString += ", cap_weld_trans_bot_g=" & IIf(IsNothing(pr.cap_weld_trans_bot_g), "Null", pr.cap_weld_trans_bot_g.ToString)
        updateString += ", cap_weld_long_bot_g=" & IIf(IsNothing(pr.cap_weld_long_bot_g), "Null", pr.cap_weld_long_bot_g.ToString)
        updateString += ", cap_weld_trans_top_g=" & IIf(IsNothing(pr.cap_weld_trans_top_g), "Null", pr.cap_weld_trans_top_g.ToString)
        updateString += ", cap_weld_long_top_g=" & IIf(IsNothing(pr.cap_weld_long_top_g), "Null", pr.cap_weld_long_top_g.ToString)
        updateString += ", cap_comp_xx_h=" & IIf(IsNothing(pr.cap_comp_xx_h), "Null", pr.cap_comp_xx_h.ToString)
        updateString += ", cap_comp_yy_h=" & IIf(IsNothing(pr.cap_comp_yy_h), "Null", pr.cap_comp_yy_h.ToString)
        updateString += ", cap_tens_yield_h=" & IIf(IsNothing(pr.cap_tens_yield_h), "Null", pr.cap_tens_yield_h.ToString)
        updateString += ", cap_tens_rupture_h=" & IIf(IsNothing(pr.cap_tens_rupture_h), "Null", pr.cap_tens_rupture_h.ToString)
        updateString += ", cap_shear_h=" & IIf(IsNothing(pr.cap_shear_h), "Null", pr.cap_shear_h.ToString)
        updateString += ", cap_bolt_shear_bot_h=" & IIf(IsNothing(pr.cap_bolt_shear_bot_h), "Null", pr.cap_bolt_shear_bot_h.ToString)
        updateString += ", cap_bolt_shear_top_h=" & IIf(IsNothing(pr.cap_bolt_shear_top_h), "Null", pr.cap_bolt_shear_top_h.ToString)
        updateString += ", cap_boltshaft_bearing_nodeform_bot_h=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_bot_h), "Null", pr.cap_boltshaft_bearing_nodeform_bot_h.ToString)
        updateString += ", cap_boltshaft_bearing_deform_bot_h=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_bot_h), "Null", pr.cap_boltshaft_bearing_deform_bot_h.ToString)
        updateString += ", cap_boltshaft_bearing_nodeform_top_h=" & IIf(IsNothing(pr.cap_boltshaft_bearing_nodeform_top_h), "Null", pr.cap_boltshaft_bearing_nodeform_top_h.ToString)
        updateString += ", cap_boltshaft_bearing_deform_top_h=" & IIf(IsNothing(pr.cap_boltshaft_bearing_deform_top_h), "Null", pr.cap_boltshaft_bearing_deform_top_h.ToString)
        updateString += ", cap_boltreinf_bearing_nodeform_bot_h=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_bot_h), "Null", pr.cap_boltreinf_bearing_nodeform_bot_h.ToString)
        updateString += ", cap_boltreinf_bearing_deform_bot_h=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_bot_h), "Null", pr.cap_boltreinf_bearing_deform_bot_h.ToString)
        updateString += ", cap_boltreinf_bearing_nodeform_top_h=" & IIf(IsNothing(pr.cap_boltreinf_bearing_nodeform_top_h), "Null", pr.cap_boltreinf_bearing_nodeform_top_h.ToString)
        updateString += ", cap_boltreinf_bearing_deform_top_h=" & IIf(IsNothing(pr.cap_boltreinf_bearing_deform_top_h), "Null", pr.cap_boltreinf_bearing_deform_top_h.ToString)
        updateString += ", cap_weld_trans_bot_h=" & IIf(IsNothing(pr.cap_weld_trans_bot_h), "Null", pr.cap_weld_trans_bot_h.ToString)
        updateString += ", cap_weld_long_bot_h=" & IIf(IsNothing(pr.cap_weld_long_bot_h), "Null", pr.cap_weld_long_bot_h.ToString)
        updateString += ", cap_weld_trans_top_h=" & IIf(IsNothing(pr.cap_weld_trans_top_h), "Null", pr.cap_weld_trans_top_h.ToString)
        updateString += ", cap_weld_long_top_h=" & IIf(IsNothing(pr.cap_weld_long_top_h), "Null", pr.cap_weld_long_top_h.ToString)
        updateString += " WHERE ID = " & pr.reinf_db_id.ToString

        Return updateString
    End Function

    Private Function UpdatePropBolt(ByVal pb As PropBolt) As String
        Dim updateString As String = ""

        updateString += "UPDATE bolt_prop_flat_plate SET "
        updateString += " bolt_db_id=" & IIf(IsNothing(pb.bolt_db_id), "Null", pb.bolt_db_id.ToString)
        updateString += ", name=" & IIf(IsNothing(pb.name), "Null", "'" & pb.name.ToString & "'")
        updateString += ", description=" & IIf(IsNothing(pb.description), "Null", "'" & pb.description.ToString & "'")
        updateString += ", diam=" & IIf(IsNothing(pb.diam), "Null", pb.diam.ToString)
        updateString += ", area=" & IIf(IsNothing(pb.area), "Null", pb.area.ToString)
        updateString += ", fu_bolt=" & IIf(IsNothing(pb.fu_bolt), "Null", pb.fu_bolt.ToString)
        updateString += ", sleeve_diam_out=" & IIf(IsNothing(pb.sleeve_diam_out), "Null", pb.sleeve_diam_out.ToString)
        updateString += ", sleeve_diam_in=" & IIf(IsNothing(pb.sleeve_diam_in), "Null", pb.sleeve_diam_in.ToString)
        updateString += ", fu_sleeve=" & IIf(IsNothing(pb.fu_sleeve), "Null", pb.fu_sleeve.ToString)
        updateString += ", bolt_n_sleeve_shear_revF=" & IIf(IsNothing(pb.bolt_n_sleeve_shear_revF), "Null", pb.bolt_n_sleeve_shear_revF.ToString)
        updateString += ", bolt_x_sleeve_shear_revF=" & IIf(IsNothing(pb.bolt_x_sleeve_shear_revF), "Null", pb.bolt_x_sleeve_shear_revF.ToString)
        updateString += ", bolt_n_sleeve_shear_revG=" & IIf(IsNothing(pb.bolt_n_sleeve_shear_revG), "Null", pb.bolt_n_sleeve_shear_revG.ToString)
        updateString += ", bolt_x_sleeve_shear_revG=" & IIf(IsNothing(pb.bolt_x_sleeve_shear_revG), "Null", pb.bolt_x_sleeve_shear_revG.ToString)
        updateString += ", bolt_n_sleeve_shear_revH=" & IIf(IsNothing(pb.bolt_n_sleeve_shear_revH), "Null", pb.bolt_n_sleeve_shear_revH.ToString)
        updateString += ", bolt_x_sleeve_shear_revH=" & IIf(IsNothing(pb.bolt_x_sleeve_shear_revH), "Null", pb.bolt_x_sleeve_shear_revH.ToString)
        updateString += ", rb_applied_revH=" & IIf(IsNothing(pb.rb_applied_revH), "Null", "'" & pb.rb_applied_revH.ToString & "'")
        updateString += " WHERE ID = " & pb.bolt_db_id.ToString

        Return updateString
    End Function

    Private Function UpdatePropMatl(ByVal pm As PropMatl) As String
        Dim updateString As String = ""

        updateString += "UPDATE matl_prop_flat_plate SET "
        updateString += " matl_db_id=" & IIf(IsNothing(pm.matl_db_id), "Null", pm.matl_db_id.ToString)
        updateString += ", name=" & IIf(IsNothing(pm.name), "Null", "'" & pm.name.ToString & "'")
        updateString += ", fy=" & IIf(IsNothing(pm.fy), "Null", pm.fy.ToString)
        updateString += ", fu=" & IIf(IsNothing(pm.fu), "Null", pm.fu.ToString)
        updateString += " WHERE ID = " & pm.matl_db_id.ToString

        Return updateString
    End Function

#End Region

#Region "General"
    Public Sub Clear()
        ExcelFilePath = ""
        Poles.Clear()

        'Remove all datatables from the main dataset
        For Each item As EXCELDTParameter In CCIpoleExcelDTParameters()
            Try
                ds.Tables.Remove(item.xlsDatatable)
            Catch ex As Exception
            End Try
        Next

        For Each item As SQLParameter In CCIpoleSQLDataTables()
            Try
                ds.Tables.Remove(item.sqlDatatable)
            Catch ex As Exception
            End Try
        Next
    End Sub

    Private Function CCIpoleSQLDataTables() As List(Of SQLParameter)
        Dim MyParameters As New List(Of SQLParameter)

        MyParameters.Add(New SQLParameter("CCIpole General SQL", "CCIpole (SELECT General).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Criteria SQL", "CCIpole (SELECT Criteria).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Pole Sections SQL", "CCIpole (SELECT Pole Sections).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Pole Reinf Sections SQL", "CCIpole (SELECT Pole Reinf Sections).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Reinf Groups SQL", "CCIpole (SELECT Reinf Groups).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Reinf Details SQL", "CCIpole (SELECT Reinf Details).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Int Groups SQL", "CCIpole (SELECT Int Groups).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Int Details SQL", "CCIpole (SELECT Int Details).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Pole Reinf Results SQL", "CCIpole (SELECT Pole Reinf Results).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Reinf Property Details SQL", "CCIpole (SELECT Prop Reinfs).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Bolt Property Details SQL", "CCIpole (SELECT Prop Bolts).sql"))
        MyParameters.Add(New SQLParameter("CCIpole Matl Property Details SQL", "CCIpole (SELECT Prop Matls).sql"))

        Return MyParameters
    End Function

    Private Function CCIpoleExcelDTParameters() As List(Of EXCELDTParameter)
        Dim MyParameters As New List(Of EXCELDTParameter)

        MyParameters.Add(New EXCELDTParameter("CCIpole General EXCEL", "A2:B3", "Analysis Criteria (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Criteria EXCEL", "B2:J3", "Analysis Criteria (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Pole Sections EXCEL", "A2:V20", "Unreinf Pole (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Pole Reinf Sections EXCEL", "A2:V200", "Reinf Pole (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Reinf Groups EXCEL", "A2:H50", "Reinf Groups (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Reinf Details EXCEL", "A2:F200", "Reinf ID (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Int Groups EXCEL", "A2:G50", "Interference Groups (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Int Details EXCEL", "A2:F200", "Interference ID (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Pole Reinf Results EXCEL", "A2:F1000", "Reinf Results (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Reinf Property Details EXCEL", "A2:DX50", "Reinforcements (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Bolt Property Details EXCEL", "A2:R20", "Bolts (SAPI)"))
        MyParameters.Add(New EXCELDTParameter("CCIpole Matl Property Details EXCEL", "A2:F20", "Materials (SAPI)"))

        Return MyParameters
    End Function
#End Region

#Region "Check Changes"

    Function CheckChanges(ByVal xlPole As CCIpole, ByVal sqlPole As CCIpole) As Boolean
        Dim changesMade As Boolean = False

        'Check Pole Analysis Criteria
        For Each pac As PoleCriteria In xlPole.criteria
            For Each sqlpac As PoleCriteria In sqlPole.criteria
                If pac.criteria_id = sqlpac.criteria_id Then
                    If Check1Change(pac.upper_structure_type, sqlpac.upper_structure_type, "CCIpole", "Upper_Structure_Type") Then changesMade = True
                    If Check1Change(pac.analysis_deg, sqlpac.analysis_deg, "CCIpole", "Analysis_Deg") Then changesMade = True
                    If Check1Change(pac.geom_increment_length, sqlpac.geom_increment_length, "CCIpole", "Geom_Increment_Length") Then changesMade = True
                    If Check1Change(pac.vnum, sqlpac.vnum, "CCIpole", "Vnum") Then changesMade = True
                    If Check1Change(pac.check_connections, sqlpac.check_connections, "CCIpole", "Check_Connections") Then changesMade = True
                    If Check1Change(pac.hole_deformation, sqlpac.hole_deformation, "CCIpole", "Hole_Deformation") Then changesMade = True
                    If Check1Change(pac.ineff_mod_check, sqlpac.ineff_mod_check, "CCIpole", "Ineff_Mod_Check") Then changesMade = True
                    If Check1Change(pac.modified, sqlpac.modified, "CCIpole", "Modified") Then changesMade = True
                    Exit For
                ElseIf pac.criteria_id = 0 Then
                    If Check1Change(pac.upper_structure_type, Nothing, "CCIpole", "Upper_Structure_Type") Then changesMade = True
                    If Check1Change(pac.analysis_deg, Nothing, "CCIpole", "Analysis_Deg") Then changesMade = True
                    If Check1Change(pac.geom_increment_length, Nothing, "CCIpole", "Geom_Increment_Length") Then changesMade = True
                    If Check1Change(pac.vnum, Nothing, "CCIpole", "Vnum") Then changesMade = True
                    If Check1Change(pac.check_connections, Nothing, "CCIpole", "Check_Connections") Then changesMade = True
                    If Check1Change(pac.hole_deformation, Nothing, "CCIpole", "Hole_Deformation") Then changesMade = True
                    If Check1Change(pac.ineff_mod_check, Nothing, "CCIpole", "Ineff_Mod_Check") Then changesMade = True
                    If Check1Change(pac.modified, Nothing, "CCIpole", "Modified") Then changesMade = True
                    Exit For
                End If
            Next
        Next

        'Check Unreinf Pole Sections
        For Each xlps As PoleSection In xlPole.unreinf_sections
            For Each sqlps As PoleSection In sqlPole.unreinf_sections
                If xlps.section_id = sqlps.section_id Then
                    If Check1Change(xlps.local_section_id, sqlps.local_section_id, "CCIpole", "Analysis_Section_Id " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.elev_bot, sqlps.elev_bot, "CCIpole", "Elev_Bot " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.elev_top, sqlps.elev_top, "CCIpole", "Elev_Top " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.length_section, sqlps.length_section, "CCIpole", "Length_Section " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.length_splice, sqlps.length_splice, "CCIpole", "Length_Splice " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.num_sides, sqlps.num_sides, "CCIpole", "Num_Sides " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.diam_bot, sqlps.diam_bot, "CCIpole", "Diam_Bot " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.diam_top, sqlps.diam_top, "CCIpole", "Diam_Top " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.wall_thickness, sqlps.wall_thickness, "CCIpole", "Wall_Thickness " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.bend_radius, sqlps.bend_radius, "CCIpole", "Bend_Radius " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.steel_grade_id, sqlps.steel_grade_id, "CCIpole", "Steel_Grade_Id " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.pole_type, sqlps.pole_type, "CCIpole", "Pole_Type " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.section_name, sqlps.section_name, "CCIpole", "Section_Name " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.socket_length, sqlps.socket_length, "CCIpole", "Socket_Length " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.weight_mult, sqlps.weight_mult, "CCIpole", "Weight_Mult " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.wp_mult, sqlps.wp_mult, "CCIpole", "Wp_Mult " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.af_factor, sqlps.af_factor, "CCIpole", "Af_Factor " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.ar_factor, sqlps.ar_factor, "CCIpole", "Ar_Factor " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.round_area_ratio, sqlps.round_area_ratio, "CCIpole", "Round_Area_Ratio " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.flat_area_ratio, sqlps.flat_area_ratio, "CCIpole", "Flat_Area_Ratio " & xlps.section_id.ToString) Then changesMade = True
                    Exit For
                ElseIf xlps.section_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them. 
                    If Check1Change(xlps.local_section_id, Nothing, "CCIpole", "Analysis_Section_Id " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.elev_bot, Nothing, "CCIpole", "Elev_Bot " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.elev_top, Nothing, "CCIpole", "Elev_Top " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.length_section, Nothing, "CCIpole", "Length_Section " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.length_splice, Nothing, "CCIpole", "Length_Splice " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.num_sides, Nothing, "CCIpole", "Num_Sides " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.diam_bot, Nothing, "CCIpole", "Diam_Bot " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.diam_top, Nothing, "CCIpole", "Diam_Top " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.wall_thickness, Nothing, "CCIpole", "Wall_Thickness " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.bend_radius, Nothing, "CCIpole", "Bend_Radius " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.steel_grade_id, Nothing, "CCIpole", "Steel_Grade_Id " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.pole_type, Nothing, "CCIpole", "Pole_Type " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.section_name, Nothing, "CCIpole", "Section_Name " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.socket_length, Nothing, "CCIpole", "Socket_Length " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.weight_mult, Nothing, "CCIpole", "Weight_Mult " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.wp_mult, Nothing, "CCIpole", "Wp_Mult " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.af_factor, Nothing, "CCIpole", "Af_Factor " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.ar_factor, Nothing, "CCIpole", "Ar_Factor " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.round_area_ratio, Nothing, "CCIpole", "Round_Area_Ratio " & xlps.section_id.ToString) Then changesMade = True
                    If Check1Change(xlps.flat_area_ratio, Nothing, "CCIpole", "Flat_Area_Ratio " & xlps.section_id.ToString) Then changesMade = True
                    Exit For
                End If
            Next
        Next

        'Check Reinf Pole Sections
        For Each xlprs As PoleReinfSection In xlPole.reinf_sections
            For Each sqlprs As PoleReinfSection In sqlPole.reinf_sections
                If xlprs.section_id = sqlprs.section_id Then
                    If Check1Change(xlprs.local_section_id, sqlprs.local_section_id, "CCIpole", "Analysis_Section_Id " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.elev_bot, sqlprs.elev_bot, "CCIpole", "Elev_Bot " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.elev_top, sqlprs.elev_top, "CCIpole", "Elev_Top " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.length_section, sqlprs.length_section, "CCIpole", "Length_Section " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.length_splice, sqlprs.length_splice, "CCIpole", "Length_Splice " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.num_sides, sqlprs.num_sides, "CCIpole", "Num_Sides " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.diam_bot, sqlprs.diam_bot, "CCIpole", "Diam_Bot " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.diam_top, sqlprs.diam_top, "CCIpole", "Diam_Top " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.wall_thickness, sqlprs.wall_thickness, "CCIpole", "Wall_Thickness " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.bend_radius, sqlprs.bend_radius, "CCIpole", "Bend_Radius " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.steel_grade_id, sqlprs.steel_grade_id, "CCIpole", "Steel_Grade_Id " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.pole_type, sqlprs.pole_type, "CCIpole", "Pole_Type " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.weight_mult, sqlprs.weight_mult, "CCIpole", "Weight_Mult " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.section_name, sqlprs.section_name, "CCIpole", "Section_Name " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.socket_length, sqlprs.socket_length, "CCIpole", "Socket_Length " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.wp_mult, sqlprs.wp_mult, "CCIpole", "Wp_Mult " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.af_factor, sqlprs.af_factor, "CCIpole", "Af_Factor " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.ar_factor, sqlprs.ar_factor, "CCIpole", "Ar_Factor " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.round_area_ratio, sqlprs.round_area_ratio, "CCIpole", "Round_Area_Ratio " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.flat_area_ratio, sqlprs.flat_area_ratio, "CCIpole", "Flat_Area_Ratio " & xlprs.section_id.ToString) Then changesMade = True
                    Exit For
                ElseIf xlprs.section_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                    If Check1Change(xlprs.local_section_id, Nothing, "CCIpole", "Analysis_Section_Id " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.elev_bot, Nothing, "CCIpole", "Elev_Bot " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.elev_top, Nothing, "CCIpole", "Elev_Top " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.length_section, Nothing, "CCIpole", "Length_Section " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.length_splice, Nothing, "CCIpole", "Length_Splice " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.num_sides, Nothing, "CCIpole", "Num_Sides " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.diam_bot, Nothing, "CCIpole", "Diam_Bot " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.diam_top, Nothing, "CCIpole", "Diam_Top " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.wall_thickness, Nothing, "CCIpole", "Wall_Thickness " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.bend_radius, Nothing, "CCIpole", "Bend_Radius " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.steel_grade_id, Nothing, "CCIpole", "Steel_Grade_Id " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.pole_type, Nothing, "CCIpole", "Pole_Type " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.weight_mult, Nothing, "CCIpole", "Weight_Mult " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.section_name, Nothing, "CCIpole", "Section_Name " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.socket_length, Nothing, "CCIpole", "Socket_Length " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.wp_mult, Nothing, "CCIpole", "Wp_Mult " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.af_factor, Nothing, "CCIpole", "Af_Factor " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.ar_factor, Nothing, "CCIpole", "Ar_Factor " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.round_area_ratio, Nothing, "CCIpole", "Round_Area_Ratio " & xlprs.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprs.flat_area_ratio, Nothing, "CCIpole", "Flat_Area_Ratio " & xlprs.section_id.ToString) Then changesMade = True
                    Exit For
                End If
            Next
        Next


        'Check Reinf Groups
        For Each xlprg As PoleReinfGroup In xlPole.reinf_groups
            For Each sqlprg As PoleReinfGroup In sqlPole.reinf_groups
                If xlprg.reinf_group_id = sqlprg.reinf_group_id Then
                    If Check1Change(xlprg.elev_bot_actual, sqlprg.elev_bot_actual, "CCIpole", "Elev_Bot_Actual " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    If Check1Change(xlprg.elev_bot_eff, sqlprg.elev_bot_eff, "CCIpole", "Elev_Bot_Eff " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    If Check1Change(xlprg.elev_top_actual, sqlprg.elev_top_actual, "CCIpole", "Elev_Top_Actual " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    If Check1Change(xlprg.elev_top_eff, sqlprg.elev_top_eff, "CCIpole", "Elev_Top_Eff " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    If Check1Change(xlprg.reinf_db_id, sqlprg.reinf_db_id, "CCIpole", "Reinf_Db_Id " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    Exit For
                ElseIf xlprg.reinf_group_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                    If Check1Change(xlprg.elev_bot_actual, Nothing, "CCIpole", "Elev_Bot_Actual " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    If Check1Change(xlprg.elev_bot_eff, Nothing, "CCIpole", "Elev_Bot_Eff " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    If Check1Change(xlprg.elev_top_actual, Nothing, "CCIpole", "Elev_Top_Actual " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    If Check1Change(xlprg.elev_top_eff, Nothing, "CCIpole", "Elev_Top_Eff " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    If Check1Change(xlprg.reinf_db_id, Nothing, "CCIpole", "Reinf_Db_Id " & xlprg.reinf_group_id.ToString) Then changesMade = True
                    Exit For
                End If

                'Check Reinf Details
                For Each xlprd As PoleReinfDetail In xlprg.reinf_ids
                    For Each sqlprd As PoleReinfDetail In sqlprg.reinf_ids
                        If xlprd.reinf_id = sqlprd.reinf_id Then
                            If Check1Change(xlprd.pole_flat, sqlprd.pole_flat, "CCIpole", "Pole_Flat " & xlprd.reinf_id.ToString) Then changesMade = True
                            If Check1Change(xlprd.horizontal_offset, sqlprd.horizontal_offset, "CCIpole", "Horizontal_Offset " & xlprd.reinf_id.ToString) Then changesMade = True
                            If Check1Change(xlprd.rotation, sqlprd.rotation, "CCIpole", "Rotation " & xlprd.reinf_id.ToString) Then changesMade = True
                            If Check1Change(xlprd.note, sqlprd.note, "CCIpole", "Note " & xlprd.reinf_id.ToString) Then changesMade = True
                            Exit For
                        ElseIf xlprd.reinf_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                            If Check1Change(xlprd.pole_flat, Nothing, "CCIpole", "Pole_Flat " & xlprd.reinf_id.ToString) Then changesMade = True
                            If Check1Change(xlprd.horizontal_offset, Nothing, "CCIpole", "Horizontal_Offset " & xlprd.reinf_id.ToString) Then changesMade = True
                            If Check1Change(xlprd.rotation, Nothing, "CCIpole", "Rotation " & xlprd.reinf_id.ToString) Then changesMade = True
                            If Check1Change(xlprd.note, Nothing, "CCIpole", "Note " & xlprd.reinf_id.ToString) Then changesMade = True
                            Exit For
                        End If
                    Next
                Next
            Next
        Next

        'Check Interference Groups
        For Each xlpig As PoleIntGroup In xlPole.int_groups
            For Each sqlpig As PoleIntGroup In sqlPole.int_groups
                If xlpig.interference_group_id = sqlpig.interference_group_id Then
                    If Check1Change(xlpig.elev_bot, sqlpig.elev_bot, "CCIpole", "Elev_Bot " & xlpig.interference_group_id.ToString) Then changesMade = True
                    If Check1Change(xlpig.elev_top, sqlpig.elev_top, "CCIpole", "Elev_Top " & xlpig.interference_group_id.ToString) Then changesMade = True
                    If Check1Change(xlpig.width, sqlpig.width, "CCIpole", "Width " & xlpig.interference_group_id.ToString) Then changesMade = True
                    If Check1Change(xlpig.description, sqlpig.description, "CCIpole", "Description " & xlpig.interference_group_id.ToString) Then changesMade = True
                    Exit For
                ElseIf xlpig.interference_group_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                    If Check1Change(xlpig.elev_bot, Nothing, "CCIpole", "Elev_Bot " & xlpig.interference_group_id.ToString) Then changesMade = True
                    If Check1Change(xlpig.elev_top, Nothing, "CCIpole", "Elev_Top " & xlpig.interference_group_id.ToString) Then changesMade = True
                    If Check1Change(xlpig.width, Nothing, "CCIpole", "Width " & xlpig.interference_group_id.ToString) Then changesMade = True
                    If Check1Change(xlpig.description, Nothing, "CCIpole", "Description " & xlpig.interference_group_id.ToString) Then changesMade = True
                    Exit For
                End If

                'Check Interference Details
                For Each xlpid As PoleIntDetail In xlpig.int_ids
                    For Each sqlpid As PoleIntDetail In sqlpig.int_ids
                        If xlpid.interference_id = sqlpid.interference_id Then
                            If Check1Change(xlpid.pole_flat, sqlpid.pole_flat, "CCIpole", "Pole_Flat " & xlpid.interference_id.ToString) Then changesMade = True
                            If Check1Change(xlpid.horizontal_offset, sqlpid.horizontal_offset, "CCIpole", "Horizontal_Offset " & xlpid.interference_id.ToString) Then changesMade = True
                            If Check1Change(xlpid.rotation, sqlpid.rotation, "CCIpole", "Rotation " & xlpid.interference_id.ToString) Then changesMade = True
                            If Check1Change(xlpid.note, sqlpid.note, "CCIpole", "Note " & xlpid.interference_id.ToString) Then changesMade = True
                            Exit For
                        ElseIf xlpid.interference_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                            If Check1Change(xlpid.pole_flat, Nothing, "CCIpole", "Pole_Flat " & xlpid.interference_id.ToString) Then changesMade = True
                            If Check1Change(xlpid.horizontal_offset, Nothing, "CCIpole", "Horizontal_Offset " & xlpid.interference_id.ToString) Then changesMade = True
                            If Check1Change(xlpid.rotation, Nothing, "CCIpole", "Rotation " & xlpid.interference_id.ToString) Then changesMade = True
                            If Check1Change(xlpid.note, Nothing, "CCIpole", "Note " & xlpid.interference_id.ToString) Then changesMade = True
                            Exit For
                        End If
                    Next
                Next
            Next
        Next






        'Check Reinf Results
        For Each xlprr As PoleReinfResults In xlPole.reinf_section_results
            For Each sqlprr As PoleReinfResults In sqlPole.reinf_section_results
                If xlprr.section_id = sqlprr.section_id Then
                    'If Check1Change(xlprr.work_order_seq_num, sqlprr.work_order_seq_num, "CCIpole", "Work_Order_Seq_Num " & xlprr.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprr.reinf_group_id, sqlprr.reinf_group_id, "CCIpole", "Reinf_Group_Id " & xlprr.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprr.result_lkup_value, sqlprr.result_lkup_value, "CCIpole", "Result_Lkup_Value " & xlprr.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprr.rating, sqlprr.rating, "CCIpole", "Rating " & xlprr.section_id.ToString) Then changesMade = True
                    Exit For
                ElseIf xlprr.section_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                    'If Check1Change(xlprr.work_order_seq_num, Nothing, "CCIpole", "Work_Order_Seq_Num " & xlprr.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprr.reinf_group_id, Nothing, "CCIpole", "Reinf_Group_Id " & xlprr.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprr.result_lkup_value, Nothing, "CCIpole", "Result_Lkup_Value " & xlprr.section_id.ToString) Then changesMade = True
                    If Check1Change(xlprr.rating, Nothing, "CCIpole", "Rating " & xlprr.section_id.ToString) Then changesMade = True
                    Exit For
                End If
            Next
        Next


        'Check Custom Reinforcement Properties
        For Each xlpr As PropReinf In xlPole.reinfs
            For Each sqlpr As PropReinf In sqlPole.reinfs
                If xlpr.reinf_db_id = sqlpr.reinf_db_id Then
                    If Check1Change(xlpr.name, sqlpr.name, "CCIpole", "Name " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.type, sqlpr.type, "CCIpole", "Type " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.b, sqlpr.b, "CCIpole", "B " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.h, sqlpr.h, "CCIpole", "H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.sr_diam, sqlpr.sr_diam, "CCIpole", "Sr_Diam " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_thkns_web, sqlpr.channel_thkns_web, "CCIpole", "Channel_Thkns_Web " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_thkns_flange, sqlpr.channel_thkns_flange, "CCIpole", "Channel_Thkns_Flange " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_eo, sqlpr.channel_eo, "CCIpole", "Channel_Eo " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_J, sqlpr.channel_J, "CCIpole", "Channel_J " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_Cw, sqlpr.channel_Cw, "CCIpole", "Channel_Cw " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.area_gross, sqlpr.area_gross, "CCIpole", "Area_Gross " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.centroid, sqlpr.centroid, "CCIpole", "Centroid " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.istension, sqlpr.istension, "CCIpole", "Istension " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.matl_id, sqlpr.matl_id, "CCIpole", "Matl_Id " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Ix, sqlpr.Ix, "CCIpole", "Ix " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Iy, sqlpr.Iy, "CCIpole", "Iy " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Lu, sqlpr.Lu, "CCIpole", "Lu " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Kx, sqlpr.Kx, "CCIpole", "Kx " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Ky, sqlpr.Ky, "CCIpole", "Ky " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_hole_size, sqlpr.bolt_hole_size, "CCIpole", "Bolt_Hole_Size " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.area_net, sqlpr.area_net, "CCIpole", "Area_Net " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.shear_lag, sqlpr.shear_lag, "CCIpole", "Shear_Lag " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_type_bot, sqlpr.connection_type_bot, "CCIpole", "Connection_Type_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revF_bot, sqlpr.connection_cap_revF_bot, "CCIpole", "Connection_Cap_Revf_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revG_bot, sqlpr.connection_cap_revG_bot, "CCIpole", "Connection_Cap_Revg_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revH_bot, sqlpr.connection_cap_revH_bot, "CCIpole", "Connection_Cap_Revh_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_id_bot, sqlpr.bolt_id_bot, "CCIpole", "bolt_id_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_N_or_X_bot, sqlpr.bolt_N_or_X_bot, "CCIpole", "Bolt_N_Or_X_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_num_bot, sqlpr.bolt_num_bot, "CCIpole", "Bolt_Num_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_spacing_bot, sqlpr.bolt_spacing_bot, "CCIpole", "Bolt_Spacing_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_edge_dist_bot, sqlpr.bolt_edge_dist_bot, "CCIpole", "Bolt_Edge_Dist_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.FlangeOrBP_connected_bot, sqlpr.FlangeOrBP_connected_bot, "CCIpole", "Flangeorbp_Connected_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_grade_bot, sqlpr.weld_grade_bot, "CCIpole", "Weld_Grade_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_type_bot, sqlpr.weld_trans_type_bot, "CCIpole", "Weld_Trans_Type_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_length_bot, sqlpr.weld_trans_length_bot, "CCIpole", "Weld_Trans_Length_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_groove_depth_bot, sqlpr.weld_groove_depth_bot, "CCIpole", "Weld_Groove_Depth_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_groove_angle_bot, sqlpr.weld_groove_angle_bot, "CCIpole", "Weld_Groove_Angle_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_fillet_size_bot, sqlpr.weld_trans_fillet_size_bot, "CCIpole", "Weld_Trans_Fillet_Size_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_eff_throat_bot, sqlpr.weld_trans_eff_throat_bot, "CCIpole", "Weld_Trans_Eff_Throat_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_type_bot, sqlpr.weld_long_type_bot, "CCIpole", "Weld_Long_Type_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_length_bot, sqlpr.weld_long_length_bot, "CCIpole", "Weld_Long_Length_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_fillet_size_bot, sqlpr.weld_long_fillet_size_bot, "CCIpole", "Weld_Long_Fillet_Size_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_eff_throat_bot, sqlpr.weld_long_eff_throat_bot, "CCIpole", "Weld_Long_Eff_Throat_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.top_bot_connections_symmetrical, sqlpr.top_bot_connections_symmetrical, "CCIpole", "Top_Bot_Connections_Symmetrical " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_type_top, sqlpr.connection_type_top, "CCIpole", "Connection_Type_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revF_top, sqlpr.connection_cap_revF_top, "CCIpole", "Connection_Cap_Revf_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revG_top, sqlpr.connection_cap_revG_top, "CCIpole", "Connection_Cap_Revg_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revH_top, sqlpr.connection_cap_revH_top, "CCIpole", "Connection_Cap_Revh_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_id_top, sqlpr.bolt_id_top, "CCIpole", "bolt_id_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_N_or_X_top, sqlpr.bolt_N_or_X_top, "CCIpole", "Bolt_N_Or_X_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_num_top, sqlpr.bolt_num_top, "CCIpole", "Bolt_Num_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_spacing_top, sqlpr.bolt_spacing_top, "CCIpole", "Bolt_Spacing_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_edge_dist_top, sqlpr.bolt_edge_dist_top, "CCIpole", "Bolt_Edge_Dist_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.FlangeOrBP_connected_top, sqlpr.FlangeOrBP_connected_top, "CCIpole", "Flangeorbp_Connected_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_grade_top, sqlpr.weld_grade_top, "CCIpole", "Weld_Grade_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_type_top, sqlpr.weld_trans_type_top, "CCIpole", "Weld_Trans_Type_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_length_top, sqlpr.weld_trans_length_top, "CCIpole", "Weld_Trans_Length_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_groove_depth_top, sqlpr.weld_groove_depth_top, "CCIpole", "Weld_Groove_Depth_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_groove_angle_top, sqlpr.weld_groove_angle_top, "CCIpole", "Weld_Groove_Angle_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_fillet_size_top, sqlpr.weld_trans_fillet_size_top, "CCIpole", "Weld_Trans_Fillet_Size_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_eff_throat_top, sqlpr.weld_trans_eff_throat_top, "CCIpole", "Weld_Trans_Eff_Throat_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_type_top, sqlpr.weld_long_type_top, "CCIpole", "Weld_Long_Type_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_length_top, sqlpr.weld_long_length_top, "CCIpole", "Weld_Long_Length_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_fillet_size_top, sqlpr.weld_long_fillet_size_top, "CCIpole", "Weld_Long_Fillet_Size_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_eff_throat_top, sqlpr.weld_long_eff_throat_top, "CCIpole", "Weld_Long_Eff_Throat_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.conn_length_bot, sqlpr.conn_length_bot, "CCIpole", "Conn_Length_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.conn_length_top, sqlpr.conn_length_top, "CCIpole", "Conn_Length_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_xx_f, sqlpr.cap_comp_xx_f, "CCIpole", "Cap_Comp_Xx_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_yy_f, sqlpr.cap_comp_yy_f, "CCIpole", "Cap_Comp_Yy_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_yield_f, sqlpr.cap_tens_yield_f, "CCIpole", "Cap_Tens_Yield_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_rupture_f, sqlpr.cap_tens_rupture_f, "CCIpole", "Cap_Tens_Rupture_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_shear_f, sqlpr.cap_shear_f, "CCIpole", "Cap_Shear_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_bot_f, sqlpr.cap_bolt_shear_bot_f, "CCIpole", "Cap_Bolt_Shear_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_top_f, sqlpr.cap_bolt_shear_top_f, "CCIpole", "Cap_Bolt_Shear_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_bot_f, sqlpr.cap_boltshaft_bearing_nodeform_bot_f, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_bot_f, sqlpr.cap_boltshaft_bearing_deform_bot_f, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_top_f, sqlpr.cap_boltshaft_bearing_nodeform_top_f, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_top_f, sqlpr.cap_boltshaft_bearing_deform_top_f, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_bot_f, sqlpr.cap_boltreinf_bearing_nodeform_bot_f, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_bot_f, sqlpr.cap_boltreinf_bearing_deform_bot_f, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_top_f, sqlpr.cap_boltreinf_bearing_nodeform_top_f, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_top_f, sqlpr.cap_boltreinf_bearing_deform_top_f, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_bot_f, sqlpr.cap_weld_trans_bot_f, "CCIpole", "Cap_Weld_Trans_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_bot_f, sqlpr.cap_weld_long_bot_f, "CCIpole", "Cap_Weld_Long_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_top_f, sqlpr.cap_weld_trans_top_f, "CCIpole", "Cap_Weld_Trans_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_top_f, sqlpr.cap_weld_long_top_f, "CCIpole", "Cap_Weld_Long_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_xx_g, sqlpr.cap_comp_xx_g, "CCIpole", "Cap_Comp_Xx_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_yy_g, sqlpr.cap_comp_yy_g, "CCIpole", "Cap_Comp_Yy_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_yield_g, sqlpr.cap_tens_yield_g, "CCIpole", "Cap_Tens_Yield_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_rupture_g, sqlpr.cap_tens_rupture_g, "CCIpole", "Cap_Tens_Rupture_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_shear_g, sqlpr.cap_shear_g, "CCIpole", "Cap_Shear_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_bot_g, sqlpr.cap_bolt_shear_bot_g, "CCIpole", "Cap_Bolt_Shear_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_top_g, sqlpr.cap_bolt_shear_top_g, "CCIpole", "Cap_Bolt_Shear_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_bot_g, sqlpr.cap_boltshaft_bearing_nodeform_bot_g, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_bot_g, sqlpr.cap_boltshaft_bearing_deform_bot_g, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_top_g, sqlpr.cap_boltshaft_bearing_nodeform_top_g, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_top_g, sqlpr.cap_boltshaft_bearing_deform_top_g, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_bot_g, sqlpr.cap_boltreinf_bearing_nodeform_bot_g, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_bot_g, sqlpr.cap_boltreinf_bearing_deform_bot_g, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_top_g, sqlpr.cap_boltreinf_bearing_nodeform_top_g, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_top_g, sqlpr.cap_boltreinf_bearing_deform_top_g, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_bot_g, sqlpr.cap_weld_trans_bot_g, "CCIpole", "Cap_Weld_Trans_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_bot_g, sqlpr.cap_weld_long_bot_g, "CCIpole", "Cap_Weld_Long_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_top_g, sqlpr.cap_weld_trans_top_g, "CCIpole", "Cap_Weld_Trans_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_top_g, sqlpr.cap_weld_long_top_g, "CCIpole", "Cap_Weld_Long_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_xx_h, sqlpr.cap_comp_xx_h, "CCIpole", "Cap_Comp_Xx_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_yy_h, sqlpr.cap_comp_yy_h, "CCIpole", "Cap_Comp_Yy_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_yield_h, sqlpr.cap_tens_yield_h, "CCIpole", "Cap_Tens_Yield_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_rupture_h, sqlpr.cap_tens_rupture_h, "CCIpole", "Cap_Tens_Rupture_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_shear_h, sqlpr.cap_shear_h, "CCIpole", "Cap_Shear_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_bot_h, sqlpr.cap_bolt_shear_bot_h, "CCIpole", "Cap_Bolt_Shear_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_top_h, sqlpr.cap_bolt_shear_top_h, "CCIpole", "Cap_Bolt_Shear_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_bot_h, sqlpr.cap_boltshaft_bearing_nodeform_bot_h, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_bot_h, sqlpr.cap_boltshaft_bearing_deform_bot_h, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_top_h, sqlpr.cap_boltshaft_bearing_nodeform_top_h, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_top_h, sqlpr.cap_boltshaft_bearing_deform_top_h, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_bot_h, sqlpr.cap_boltreinf_bearing_nodeform_bot_h, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_bot_h, sqlpr.cap_boltreinf_bearing_deform_bot_h, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_top_h, sqlpr.cap_boltreinf_bearing_nodeform_top_h, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_top_h, sqlpr.cap_boltreinf_bearing_deform_top_h, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_bot_h, sqlpr.cap_weld_trans_bot_h, "CCIpole", "Cap_Weld_Trans_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_bot_h, sqlpr.cap_weld_long_bot_h, "CCIpole", "Cap_Weld_Long_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_top_h, sqlpr.cap_weld_trans_top_h, "CCIpole", "Cap_Weld_Trans_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_top_h, sqlpr.cap_weld_long_top_h, "CCIpole", "Cap_Weld_Long_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    Exit For
                ElseIf xlpr.reinf_db_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                    If Check1Change(xlpr.name, Nothing, "CCIpole", "Name " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.type, Nothing, "CCIpole", "Type " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.b, Nothing, "CCIpole", "B " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.h, Nothing, "CCIpole", "H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.sr_diam, Nothing, "CCIpole", "Sr_Diam " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_thkns_web, Nothing, "CCIpole", "Channel_Thkns_Web " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_thkns_flange, Nothing, "CCIpole", "Channel_Thkns_Flange " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_eo, Nothing, "CCIpole", "Channel_Eo " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_J, Nothing, "CCIpole", "Channel_J " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.channel_Cw, Nothing, "CCIpole", "Channel_Cw " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.area_gross, Nothing, "CCIpole", "Area_Gross " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.centroid, Nothing, "CCIpole", "Centroid " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.istension, Nothing, "CCIpole", "Istension " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.matl_id, Nothing, "CCIpole", "Matl_Id " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Ix, Nothing, "CCIpole", "Ix " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Iy, Nothing, "CCIpole", "Iy " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Lu, Nothing, "CCIpole", "Lu " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Kx, Nothing, "CCIpole", "Kx " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.Ky, Nothing, "CCIpole", "Ky " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_hole_size, Nothing, "CCIpole", "Bolt_Hole_Size " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.area_net, Nothing, "CCIpole", "Area_Net " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.shear_lag, Nothing, "CCIpole", "Shear_Lag " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_type_bot, Nothing, "CCIpole", "Connection_Type_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revF_bot, Nothing, "CCIpole", "Connection_Cap_Revf_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revG_bot, Nothing, "CCIpole", "Connection_Cap_Revg_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revH_bot, Nothing, "CCIpole", "Connection_Cap_Revh_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_id_bot, Nothing, "CCIpole", "bolt_id_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_N_or_X_bot, Nothing, "CCIpole", "Bolt_N_Or_X_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_num_bot, Nothing, "CCIpole", "Bolt_Num_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_spacing_bot, Nothing, "CCIpole", "Bolt_Spacing_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_edge_dist_bot, Nothing, "CCIpole", "Bolt_Edge_Dist_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.FlangeOrBP_connected_bot, Nothing, "CCIpole", "Flangeorbp_Connected_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_grade_bot, Nothing, "CCIpole", "Weld_Grade_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_type_bot, Nothing, "CCIpole", "Weld_Trans_Type_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_length_bot, Nothing, "CCIpole", "Weld_Trans_Length_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_groove_depth_bot, Nothing, "CCIpole", "Weld_Groove_Depth_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_groove_angle_bot, Nothing, "CCIpole", "Weld_Groove_Angle_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_fillet_size_bot, Nothing, "CCIpole", "Weld_Trans_Fillet_Size_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_eff_throat_bot, Nothing, "CCIpole", "Weld_Trans_Eff_Throat_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_type_bot, Nothing, "CCIpole", "Weld_Long_Type_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_length_bot, Nothing, "CCIpole", "Weld_Long_Length_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_fillet_size_bot, Nothing, "CCIpole", "Weld_Long_Fillet_Size_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_eff_throat_bot, Nothing, "CCIpole", "Weld_Long_Eff_Throat_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.top_bot_connections_symmetrical, Nothing, "CCIpole", "Top_Bot_Connections_Symmetrical " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_type_top, Nothing, "CCIpole", "Connection_Type_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revF_top, Nothing, "CCIpole", "Connection_Cap_Revf_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revG_top, Nothing, "CCIpole", "Connection_Cap_Revg_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.connection_cap_revH_top, Nothing, "CCIpole", "Connection_Cap_Revh_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_id_top, Nothing, "CCIpole", "bolt_id_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_N_or_X_top, Nothing, "CCIpole", "Bolt_N_Or_X_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_num_top, Nothing, "CCIpole", "Bolt_Num_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_spacing_top, Nothing, "CCIpole", "Bolt_Spacing_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.bolt_edge_dist_top, Nothing, "CCIpole", "Bolt_Edge_Dist_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.FlangeOrBP_connected_top, Nothing, "CCIpole", "Flangeorbp_Connected_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_grade_top, Nothing, "CCIpole", "Weld_Grade_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_type_top, Nothing, "CCIpole", "Weld_Trans_Type_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_length_top, Nothing, "CCIpole", "Weld_Trans_Length_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_groove_depth_top, Nothing, "CCIpole", "Weld_Groove_Depth_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_groove_angle_top, Nothing, "CCIpole", "Weld_Groove_Angle_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_fillet_size_top, Nothing, "CCIpole", "Weld_Trans_Fillet_Size_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_trans_eff_throat_top, Nothing, "CCIpole", "Weld_Trans_Eff_Throat_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_type_top, Nothing, "CCIpole", "Weld_Long_Type_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_length_top, Nothing, "CCIpole", "Weld_Long_Length_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_fillet_size_top, Nothing, "CCIpole", "Weld_Long_Fillet_Size_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.weld_long_eff_throat_top, Nothing, "CCIpole", "Weld_Long_Eff_Throat_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.conn_length_bot, Nothing, "CCIpole", "Conn_Length_Bot " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.conn_length_top, Nothing, "CCIpole", "Conn_Length_Top " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_xx_f, Nothing, "CCIpole", "Cap_Comp_Xx_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_yy_f, Nothing, "CCIpole", "Cap_Comp_Yy_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_yield_f, Nothing, "CCIpole", "Cap_Tens_Yield_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_rupture_f, Nothing, "CCIpole", "Cap_Tens_Rupture_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_shear_f, Nothing, "CCIpole", "Cap_Shear_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_bot_f, Nothing, "CCIpole", "Cap_Bolt_Shear_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_top_f, Nothing, "CCIpole", "Cap_Bolt_Shear_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_bot_f, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_bot_f, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_top_f, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_top_f, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_bot_f, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_bot_f, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_top_f, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_top_f, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_bot_f, Nothing, "CCIpole", "Cap_Weld_Trans_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_bot_f, Nothing, "CCIpole", "Cap_Weld_Long_Bot_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_top_f, Nothing, "CCIpole", "Cap_Weld_Trans_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_top_f, Nothing, "CCIpole", "Cap_Weld_Long_Top_F " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_xx_g, Nothing, "CCIpole", "Cap_Comp_Xx_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_yy_g, Nothing, "CCIpole", "Cap_Comp_Yy_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_yield_g, Nothing, "CCIpole", "Cap_Tens_Yield_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_rupture_g, Nothing, "CCIpole", "Cap_Tens_Rupture_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_shear_g, Nothing, "CCIpole", "Cap_Shear_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_bot_g, Nothing, "CCIpole", "Cap_Bolt_Shear_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_top_g, Nothing, "CCIpole", "Cap_Bolt_Shear_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_bot_g, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_bot_g, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_top_g, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_top_g, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_bot_g, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_bot_g, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_top_g, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_top_g, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_bot_g, Nothing, "CCIpole", "Cap_Weld_Trans_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_bot_g, Nothing, "CCIpole", "Cap_Weld_Long_Bot_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_top_g, Nothing, "CCIpole", "Cap_Weld_Trans_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_top_g, Nothing, "CCIpole", "Cap_Weld_Long_Top_G " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_xx_h, Nothing, "CCIpole", "Cap_Comp_Xx_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_comp_yy_h, Nothing, "CCIpole", "Cap_Comp_Yy_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_yield_h, Nothing, "CCIpole", "Cap_Tens_Yield_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_tens_rupture_h, Nothing, "CCIpole", "Cap_Tens_Rupture_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_shear_h, Nothing, "CCIpole", "Cap_Shear_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_bot_h, Nothing, "CCIpole", "Cap_Bolt_Shear_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_bolt_shear_top_h, Nothing, "CCIpole", "Cap_Bolt_Shear_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_bot_h, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_bot_h, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_nodeform_top_h, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Nodeform_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltshaft_bearing_deform_top_h, Nothing, "CCIpole", "Cap_Boltshaft_Bearing_Deform_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_bot_h, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_bot_h, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_nodeform_top_h, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Nodeform_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_boltreinf_bearing_deform_top_h, Nothing, "CCIpole", "Cap_Boltreinf_Bearing_Deform_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_bot_h, Nothing, "CCIpole", "Cap_Weld_Trans_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_bot_h, Nothing, "CCIpole", "Cap_Weld_Long_Bot_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_trans_top_h, Nothing, "CCIpole", "Cap_Weld_Trans_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpr.cap_weld_long_top_h, Nothing, "CCIpole", "Cap_Weld_Long_Top_H " & xlpr.reinf_db_id.ToString) Then changesMade = True
                    Exit For
                End If
            Next
        Next


        'Check Custom Bolt Properties
        For Each xlpb As PropBolt In xlPole.bolts
            For Each sqlpb As PropBolt In sqlPole.bolts
                If xlpb.bolt_db_id = sqlpb.bolt_db_id Then
                    If Check1Change(xlpb.name, sqlpb.name, "CCIpole", "Name " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.description, sqlpb.description, "CCIpole", "Description " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.diam, sqlpb.diam, "CCIpole", "Diam " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.area, sqlpb.area, "CCIpole", "Area " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.fu_bolt, sqlpb.fu_bolt, "CCIpole", "Fu_Bolt " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.sleeve_diam_out, sqlpb.sleeve_diam_out, "CCIpole", "Sleeve_Diam_Out " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.sleeve_diam_in, sqlpb.sleeve_diam_in, "CCIpole", "Sleeve_Diam_In " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.fu_sleeve, sqlpb.fu_sleeve, "CCIpole", "Fu_Sleeve " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_n_sleeve_shear_revF, sqlpb.bolt_n_sleeve_shear_revF, "CCIpole", "Bolt_N_Sleeve_Shear_Revf " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_x_sleeve_shear_revF, sqlpb.bolt_x_sleeve_shear_revF, "CCIpole", "Bolt_X_Sleeve_Shear_Revf " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_n_sleeve_shear_revG, sqlpb.bolt_n_sleeve_shear_revG, "CCIpole", "Bolt_N_Sleeve_Shear_Revg " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_x_sleeve_shear_revG, sqlpb.bolt_x_sleeve_shear_revG, "CCIpole", "Bolt_X_Sleeve_Shear_Revg " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_n_sleeve_shear_revH, sqlpb.bolt_n_sleeve_shear_revH, "CCIpole", "Bolt_N_Sleeve_Shear_Revh " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_x_sleeve_shear_revH, sqlpb.bolt_x_sleeve_shear_revH, "CCIpole", "Bolt_X_Sleeve_Shear_Revh " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.rb_applied_revH, sqlpb.rb_applied_revH, "CCIpole", "Rb_Applied_Revh " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    Exit For
                ElseIf xlpb.bolt_db_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                    If Check1Change(xlpb.name, Nothing, "CCIpole", "Name " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.description, Nothing, "CCIpole", "Description " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.diam, Nothing, "CCIpole", "Diam " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.area, Nothing, "CCIpole", "Area " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.fu_bolt, Nothing, "CCIpole", "Fu_Bolt " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.sleeve_diam_out, Nothing, "CCIpole", "Sleeve_Diam_Out " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.sleeve_diam_in, Nothing, "CCIpole", "Sleeve_Diam_In " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.fu_sleeve, Nothing, "CCIpole", "Fu_Sleeve " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_n_sleeve_shear_revF, Nothing, "CCIpole", "Bolt_N_Sleeve_Shear_Revf " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_x_sleeve_shear_revF, Nothing, "CCIpole", "Bolt_X_Sleeve_Shear_Revf " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_n_sleeve_shear_revG, Nothing, "CCIpole", "Bolt_N_Sleeve_Shear_Revg " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_x_sleeve_shear_revG, Nothing, "CCIpole", "Bolt_X_Sleeve_Shear_Revg " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_n_sleeve_shear_revH, Nothing, "CCIpole", "Bolt_N_Sleeve_Shear_Revh " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.bolt_x_sleeve_shear_revH, Nothing, "CCIpole", "Bolt_X_Sleeve_Shear_Revh " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpb.rb_applied_revH, Nothing, "CCIpole", "Rb_Applied_Revh " & xlpb.bolt_db_id.ToString) Then changesMade = True
                    Exit For
                End If
            Next
        Next


        'Check Custom Material Properties
        For Each xlpm As PropMatl In xlPole.matls
            For Each sqlpm As PropMatl In sqlPole.matls
                If xlpm.matl_db_id = sqlpm.matl_db_id Then
                    If Check1Change(xlpm.name, sqlpm.name, "CCIpole", "Name " & xlpm.matl_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpm.fy, sqlpm.fy, "CCIpole", "Fy " & xlpm.matl_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpm.fu, sqlpm.fu, "CCIpole", "Fu " & xlpm.matl_db_id.ToString) Then changesMade = True
                    Exit For
                ElseIf xlpm.matl_db_id = 0 Then 'accounts for inserting new rows. additional rows won't have an ID associated to them.
                    If Check1Change(xlpm.name, Nothing, "CCIpole", "Name " & xlpm.matl_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpm.fy, Nothing, "CCIpole", "Fy " & xlpm.matl_db_id.ToString) Then changesMade = True
                    If Check1Change(xlpm.fu, Nothing, "CCIpole", "Fu " & xlpm.matl_db_id.ToString) Then changesMade = True
                    Exit For
                End If
            Next
        Next


        CreateChangeSummary(changeDt) 'possible alternative to listing change summary
        Return changesMade
    End Function

#End Region

End Class