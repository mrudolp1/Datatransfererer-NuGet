Option Strict On

Imports System.ComponentModel
Imports DevExpress.Spreadsheet
Imports System.Runtime.Serialization

''Adding fields to Seismic
''1. Excel
''  -Add field to Details (Sapi) tab
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
''***Code change note***
''  -requires updating Save to Excel section (search code_change) and Excel VBA z_EDS_Connection module
''  -Goal is to run USGS with default values within the applicable code


<DataContractAttribute()>
Partial Public Class CCISeismic
    Inherits EDSExcelObject

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "CCISeismic"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "load.seismic"
        End Get
    End Property
    Public Overrides ReadOnly Property TemplatePath As String
        Get
            Return IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "CCISeismic.xlsm")
        End Get
    End Property
    Public Overrides ReadOnly Property Template As Byte()
        Get
            Return CCI_Engineering_Templates.My.Resources.CCISeismic
        End Get
    End Property
    Public Overrides ReadOnly Property ExcelDTParams As List(Of EXCELDTParameter)
        'Add additional sub table references here. Table names should be consistent with EDS table names. 
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("Seismic Details", "A1:AE2", "Details (SAPI)")}

            'note: Excel table names are consistent with EDS table names to limit work required within constructors

        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String

        If _Insert = "" Then
            _Insert = CCI_Engineering_Templates.My.Resources.CCISeismic_INSERT
        End If
        SQLInsert = _Insert

        'Top Level
        SQLInsert = SQLInsert.Replace("[SEISMIC VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[SEISMIC FIELDS]", Me.SQLInsertFields)

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String
        'This section not only needs to call update commands but also needs to call insert and delete commands since subtables may involve adding or deleting records

        If _Update = "" Then
            _Update = CCI_Engineering_Templates.My.Resources.CCISeismic_UPDATE
        End If
        SQLUpdate = _Update

        'Top Level
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'Top Level
        If _Delete = "" Then
            _Delete = CCI_Engineering_Templates.My.Resources.CCISeismic_DELETE
        End If
        SQLDelete = _Delete
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)

        Return SQLDelete

    End Function

#End Region

#Region "Define"

    Private _lat_sign As String
    Private _lat_deg As Integer?
    Private _lat_min As Integer?
    Private _lat_sec As Double?
    Private _long_sign As String
    Private _long_deg As Integer?
    Private _long_min As Integer?
    Private _long_sec As Double?
    Private _use_asce As Boolean?
    Private _site_soil As String
    Private _risk_category As String
    Private _ss As Double?
    Private _s1 As Double?
    Private _tl As Double?
    Private _importance_factor_override As Boolean?
    Private _importance_factor_user As Double?
    Private _response_accel_override As Boolean?
    Private _sds_user As Double?
    Private _sd1_user As Double?
    Private _amp_factor As Double?
    Private _tia_approx_period As Boolean?
    Private _fundamental_period_user As Double?
    Private _mp_density_override As Boolean?
    Private _density_tower_material As Double?
    Private _elasticity As Double?
    Private _create_seismic_loads As Boolean?
    Private _user_force_appurtenance As Boolean?
    Private _sdc As String
    Private _design_code As String

    <Category("Seismic"), Description(""), DisplayName("Lat Sign")>
     <DataMember()> Public Property lat_sign() As String
        Get
            Return Me._lat_sign
        End Get
        Set
            Me._lat_sign = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Lat Deg")>
     <DataMember()> Public Property lat_deg() As Integer?
        Get
            Return Me._lat_deg
        End Get
        Set
            Me._lat_deg = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Lat Min")>
     <DataMember()> Public Property lat_min() As Integer?
        Get
            Return Me._lat_min
        End Get
        Set
            Me._lat_min = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Lat Sec")>
     <DataMember()> Public Property lat_sec() As Double?
        Get
            Return Me._lat_sec
        End Get
        Set
            Me._lat_sec = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Long Sign")>
     <DataMember()> Public Property long_sign() As String
        Get
            Return Me._long_sign
        End Get
        Set
            Me._long_sign = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Long Deg")>
     <DataMember()> Public Property long_deg() As Integer?
        Get
            Return Me._long_deg
        End Get
        Set
            Me._long_deg = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Long Min")>
     <DataMember()> Public Property long_min() As Integer?
        Get
            Return Me._long_min
        End Get
        Set
            Me._long_min = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Long Sec")>
     <DataMember()> Public Property long_sec() As Double?
        Get
            Return Me._long_sec
        End Get
        Set
            Me._long_sec = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Use Asce")>
     <DataMember()> Public Property use_asce() As Boolean?
        Get
            Return Me._use_asce
        End Get
        Set
            Me._use_asce = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Site Soil")>
     <DataMember()> Public Property site_soil() As String
        Get
            Return Me._site_soil
        End Get
        Set
            Me._site_soil = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Risk Category")>
     <DataMember()> Public Property risk_category() As String
        Get
            Return Me._risk_category
        End Get
        Set
            Me._risk_category = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Ss")>
     <DataMember()> Public Property ss() As Double?
        Get
            Return Me._ss
        End Get
        Set
            Me._ss = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("S1")>
     <DataMember()> Public Property s1() As Double?
        Get
            Return Me._s1
        End Get
        Set
            Me._s1 = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Tl")>
     <DataMember()> Public Property tl() As Double?
        Get
            Return Me._tl
        End Get
        Set
            Me._tl = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Importance Factor Override")>
     <DataMember()> Public Property importance_factor_override() As Boolean?
        Get
            Return Me._importance_factor_override
        End Get
        Set
            Me._importance_factor_override = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Importance Factor User")>
     <DataMember()> Public Property importance_factor_user() As Double?
        Get
            Return Me._importance_factor_user
        End Get
        Set
            Me._importance_factor_user = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Response Accel Override")>
     <DataMember()> Public Property response_accel_override() As Boolean?
        Get
            Return Me._response_accel_override
        End Get
        Set
            Me._response_accel_override = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Sds User")>
     <DataMember()> Public Property sds_user() As Double?
        Get
            Return Me._sds_user
        End Get
        Set
            Me._sds_user = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Sd1 User")>
     <DataMember()> Public Property sd1_user() As Double?
        Get
            Return Me._sd1_user
        End Get
        Set
            Me._sd1_user = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Amp Factor")>
     <DataMember()> Public Property amp_factor() As Double?
        Get
            Return Me._amp_factor
        End Get
        Set
            Me._amp_factor = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Tia Approx Period")>
     <DataMember()> Public Property tia_approx_period() As Boolean?
        Get
            Return Me._tia_approx_period
        End Get
        Set
            Me._tia_approx_period = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Fundamental Period User")>
     <DataMember()> Public Property fundamental_period_user() As Double?
        Get
            Return Me._fundamental_period_user
        End Get
        Set
            Me._fundamental_period_user = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("MP Density Override")>
     <DataMember()> Public Property mp_density_override() As Boolean?
        Get
            Return Me._mp_density_override
        End Get
        Set
            Me._mp_density_override = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Density Tower Material")>
     <DataMember()> Public Property density_tower_material() As Double?
        Get
            Return Me._density_tower_material
        End Get
        Set
            Me._density_tower_material = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Elasticity")>
     <DataMember()> Public Property elasticity() As Double?
        Get
            Return Me._elasticity
        End Get
        Set
            Me._elasticity = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Create Seismic Loads")>
     <DataMember()> Public Property create_seismic_loads() As Boolean?
        Get
            Return Me._create_seismic_loads
        End Get
        Set
            Me._create_seismic_loads = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("User Force Appurtenance")>
     <DataMember()> Public Property user_force_appurtenance() As Boolean?
        Get
            Return Me._user_force_appurtenance
        End Get
        Set
            Me._user_force_appurtenance = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Seismic Design Category")>
    <DataMember()> Public Property sdc() As String
        Get
            Return Me._sdc
        End Get
        Set
            Me._sdc = Value
        End Set
    End Property
    <Category("Seismic"), Description(""), DisplayName("Design Code")>
    <DataMember()> Public Property design_code() As String
        Get
            Return Me._design_code
        End Get
        Set
            Me._design_code = Value
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

    End Sub 'Generate Seismic from EDS

    Public Sub New(ExcelFilePath As String, Optional ByRef Parent As EDSObject = Nothing)
        Me.WorkBookPath = ExcelFilePath
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        LoadFromExcel()

    End Sub 'Generate Leg Reinforcement from Excel

    Private Sub BuildFromDataset(ByVal dr As DataRow, ByRef ds As DataSet, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing)
        'Dataset is pulled in from either EDS or Excel. True = EDS, False = Excel
        'If Parent IsNot Nothing Then Me.Absorb(Parent) 'Do not double absorb!!!

        'Not sure this is necessary, could just read the values from the structure code criteria when creating the Excel sheet (Added to Save to Excel Section)
        'Me.tia_current = Me.ParentStructure?.structureCodeCriteria?.tia_current
        'Me.rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
        'Me.seismic_design_category = Me.ParentStructure?.structureCodeCriteria?.seismic_design_category

        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.Version = DBtoStr(dr.Item("tool_version"))
        Me.bus_unit = If(EDStruefalse, DBtoStr(dr.Item("bus_unit")), Me.bus_unit) 'Not provided in Excel
        Me.structure_id = If(EDStruefalse, DBtoStr(dr.Item("structure_id")), Me.structure_id) 'Not provided in Excel
        Me.modified_person_id = If(EDStruefalse, DBtoNullableInt(dr.Item("modified_person_id")), Me.modified_person_id) 'Not provided in Excel
        Me.process_stage = If(EDStruefalse, DBtoStr(dr.Item("process_stage")), Me.process_stage) 'Not provided in Excel
        Me.lat_sign = DBtoStr(dr.Item("lat_sign"))
        Me.lat_deg = DBtoNullableInt(dr.Item("lat_deg"))
        Me.lat_min = DBtoNullableInt(dr.Item("lat_min"))
        Me.lat_sec = DBtoNullableDbl(dr.Item("lat_sec"))
        Me.long_sign = DBtoStr(dr.Item("long_sign"))
        Me.long_deg = DBtoNullableInt(dr.Item("long_deg"))
        Me.long_min = DBtoNullableInt(dr.Item("long_min"))
        Me.long_sec = DBtoNullableDbl(dr.Item("long_sec"))
        Me.use_asce = DBtoNullableBool(dr.Item("use_asce"))
        Me.site_soil = DBtoStr(dr.Item("site_soil"))
        Me.risk_category = DBtoStr(dr.Item("risk_category"))
        Me.ss = DBtoNullableDbl(dr.Item("ss"))
        Me.s1 = DBtoNullableDbl(dr.Item("s1"))
        Me.tl = DBtoNullableDbl(dr.Item("tl"))
        Me.importance_factor_override = DBtoNullableBool(dr.Item("importance_factor_override"))
        Me.importance_factor_user = DBtoNullableDbl(dr.Item("importance_factor_user"))
        Me.response_accel_override = DBtoNullableBool(dr.Item("response_accel_override"))
        Me.sds_user = DBtoNullableDbl(dr.Item("sds_user"))
        Me.sd1_user = DBtoNullableDbl(dr.Item("sd1_user"))
        Me.amp_factor = DBtoNullableDbl(dr.Item("amp_factor"))
        Me.tia_approx_period = DBtoNullableBool(dr.Item("tia_approx_period"))
        Me.fundamental_period_user = DBtoNullableDbl(dr.Item("fundamental_period_user"))
        Me.mp_density_override = DBtoNullableBool(dr.Item("mp_density_override"))
        Me.density_tower_material = DBtoNullableDbl(dr.Item("density_tower_material"))
        Me.elasticity = DBtoNullableDbl(dr.Item("elasticity"))
        Me.create_seismic_loads = DBtoNullableBool(dr.Item("create_seismic_loads"))
        Me.user_force_appurtenance = DBtoNullableBool(dr.Item("user_force_appurtenance"))
        Try
            Me.sdc = DBtoStr(dr.Item("sdc"))
        Catch ex As Exception
        End Try
        Me.design_code = DBtoStr(dr.Item("design_code"))

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

        If excelDS.Tables.Contains("Seismic Details") Then
            Dim dr = excelDS.Tables("Seismic Details").Rows(0)

            'Following used to create dataset, regardless if source was EDS or Excel. Boolean used to identify source. Excel = False
            BuildFromDataset(dr, excelDS, False, Me)

        End If
    End Sub
#End Region

#Region "Save to Excel"

    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        '''''Customize for each excel tool'''''
        Dim code_change As Boolean = False
        Dim use_asce_d As Boolean
        Dim site_soil_d, risk_category_d As String

        'Site Code Criteria
        Dim tia_current As String
        Dim site_app, site_rev As Integer?
        Dim site_lat_deg_dec, site_abs_lat_deg_dec, site_lat_deg, site_lat_min_dec, site_lat_min, site_lat_sec_dec As Double?
        Dim site_long_deg_dec, site_abs_long_deg_dec, site_long_deg, site_long_min_dec, site_long_min, site_long_sec_dec As Double?

        With wb
            'Site Code Criteria

            'Work Order
            If Not IsNothing(Me.ParentStructure?.work_order_seq_num) Then
                work_order_seq_num = Me.ParentStructure?.work_order_seq_num
                .Worksheets("Site SDC Data").Range("wo").Value = CType(work_order_seq_num, Integer)
            End If

            'App ID
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.eng_app_id) Then
                site_app = Me.ParentStructure?.structureCodeCriteria?.eng_app_id
                .Worksheets("Site SDC Data").Range("app").Value = CType(site_app, String)
            End If

            'Revision #
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.eng_app_id_revision) Then
                site_rev = Me.ParentStructure?.structureCodeCriteria?.eng_app_id_revision
                .Worksheets("Site SDC Data").Range("rev").Value = CType(site_rev, String)
            End If

            'TIA Revision- Defaulting to Rev. H-1 if not available.
            If MyTIA() = "F" Then
                tia_current = "ASCE 7-10"
            ElseIf MyTIA() = "G" Then
                tia_current = "TIA-222-G"
            ElseIf MyTIA() = "H" Then
                tia_current = "TIA-222-H-1"
            Else
                tia_current = "TIA-222-H-1"
            End If
            .Worksheets("Site SDC Data").Range("dcode").Value = CType(tia_current, String)

            'Check for code change. Code change will invalidate Site Soil and Risk Category and seismic values (Ss, S1 and TL) pulled in from EDS. 
            'When code change occurs, going to assume default values to rerun USGS and continue analysis
            If IsSomethingString(design_code) Then
                If design_code <> tia_current Then
                    code_change = True
                    .Worksheets("Details (SAPI)").Range("A4").Value = CType(True, Boolean)
                    'Determine default values
                    If tia_current = "TIA-222-H-1" Then
                        use_asce_d = False
                        site_soil_d = "D (Default)"
                        risk_category_d = "II"
                    ElseIf tia_current = "TIA-222-G" Or tia_current = "ASCE 7-10" Then
                        use_asce_d = False
                        site_soil_d = "D"
                        risk_category_d = "II"
                    End If
                End If
            End If

            'Latitude
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.lat_dec) Then
                'sign
                If Me.ParentStructure?.structureCodeCriteria?.lat_dec > 0 Then
                    .Worksheets("Site SDC Data").Range("latsign").Value = CType("+", String)
                Else
                    .Worksheets("Site SDC Data").Range("latsign").Value = CType("-", String)
                End If
                'GetValueOrDefault added to convert from double? to double. required for math
                'Degree
                site_lat_deg_dec = Me.ParentStructure?.structureCodeCriteria?.lat_dec
                site_abs_lat_deg_dec = Math.Abs(site_lat_deg_dec.GetValueOrDefault())
                site_lat_deg = Math.Floor(Math.Abs(site_lat_deg_dec.GetValueOrDefault()))
                .Worksheets("Site SDC Data").Range("latdeg").Value = CType(site_lat_deg, Integer)
                'Minute
                site_lat_min_dec = (site_abs_lat_deg_dec - site_lat_deg) * 60
                site_lat_min = Math.Floor(Math.Abs(site_lat_min_dec.GetValueOrDefault()))
                .Worksheets("Site SDC Data").Range("latmin").Value = CType(site_lat_min, Integer)
                'Second
                site_lat_sec_dec = (site_lat_min_dec - site_lat_min) * 60
                .Worksheets("Site SDC Data").Range("latsec").Value = CType(site_lat_sec_dec, Double)
            End If

            'Longitude
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.long_dec) Then
                'sign
                If Me.ParentStructure?.structureCodeCriteria?.long_dec > 0 Then
                    .Worksheets("Site SDC Data").Range("longsign").Value = CType("+", String)
                Else
                    .Worksheets("Site SDC Data").Range("longsign").Value = CType("-", String)
                End If
                'GetValueOrDefault added to convert from double? to double. required for math
                'Degree
                site_long_deg_dec = Me.ParentStructure?.structureCodeCriteria?.long_dec
                site_abs_long_deg_dec = Math.Abs(site_long_deg_dec.GetValueOrDefault())
                site_long_deg = Math.Floor(Math.Abs(site_long_deg_dec.GetValueOrDefault()))
                .Worksheets("Site SDC Data").Range("longdeg").Value = CType(site_long_deg, Integer)
                'Minute
                site_long_min_dec = (site_abs_long_deg_dec - site_long_deg) * 60
                site_long_min = Math.Floor(Math.Abs(site_long_min_dec.GetValueOrDefault()))
                .Worksheets("Site SDC Data").Range("longmin").Value = CType(site_long_min, Integer)
                'Second
                site_long_sec_dec = (site_long_min_dec - site_long_min) * 60
                .Worksheets("Site SDC Data").Range("longsec").Value = CType(site_long_sec_dec, Double)
            End If

            .Worksheets("Details (SAPI)").Range("A3").Value = CType(True, Boolean) 'Flags if sheet was last touched by EDS. If true, worksheet change event upon opening tool. 

            .Worksheets("Details (SAPI)").Range("ID").Value = CType(Me.ID, Integer)

            If Not IsNothing(Me.bus_unit) Then
                .Worksheets("Site SDC Data").Range("bu").Value = CType(Me.bus_unit, Integer)
            Else
                .Worksheets("Site SDC Data").Range("bu").ClearContents
            End If
            If Not IsNothing(Me.work_order_seq_num) Then
                .Worksheets("Site SDC Data").Range("wo").Value = CType(Me.work_order_seq_num, Integer)
            Else
                .Worksheets("Site SDC Data").Range("wo").ClearContents
            End If
            If Not IsNothing(Me.structure_id) Then
                .Worksheets("Site SDC Data").Range("strc").Value = CType(Me.structure_id, String)
            End If
            'If Not IsNothing(Me.tool_version) Then
            '    .Worksheets("Revision History").Range("Revision").Value = CType(Me.tool_version, String)
            'End If
            'If Not IsNothing(Me.lat_sign) Then
            '    .Worksheets("Site SDC Data").Range("latsign").Value = CType(Me.lat_sign, String)
            'End If
            'If Not IsNothing(Me.lat_deg) Then
            '    .Worksheets("Site SDC Data").Range("latdeg").Value = CType(Me.lat_deg, Integer)
            'Else
            '    .Worksheets("Site SDC Data").Range("latdeg").ClearContents
            'End If
            'If Not IsNothing(Me.lat_min) Then
            '    .Worksheets("Site SDC Data").Range("latmin").Value = CType(Me.lat_min, Integer)
            'Else
            '    .Worksheets("Site SDC Data").Range("latmin").ClearContents
            'End If
            'If Not IsNothing(Me.lat_sec) Then
            '    .Worksheets("Site SDC Data").Range("latsec").Value = CType(Me.lat_sec, Double)
            'Else
            '    .Worksheets("Site SDC Data").Range("latsec").ClearContents
            'End If
            'If Not IsNothing(Me.long_sign) Then
            '    .Worksheets("Site SDC Data").Range("longsign").Value = CType(Me.long_sign, String)
            'End If
            'If Not IsNothing(Me.long_deg) Then
            '    .Worksheets("Site SDC Data").Range("longdeg").Value = CType(Me.long_deg, Integer)
            'Else
            '    .Worksheets("Site SDC Data").Range("longdeg").ClearContents
            'End If
            'If Not IsNothing(Me.long_min) Then
            '    .Worksheets("Site SDC Data").Range("longmin").Value = CType(Me.long_min, Integer)
            'Else
            '    .Worksheets("Site SDC Data").Range("longmin").ClearContents
            'End If
            'If Not IsNothing(Me.long_sec) Then
            '    .Worksheets("Site SDC Data").Range("longsec").Value = CType(Me.long_sec, Double)
            'Else
            '    .Worksheets("Site SDC Data").Range("longsec").ClearContents
            'End If
            If code_change Then
                .Worksheets("Reference").Range("Use_ASCE").Value = CType(use_asce_d, Boolean)
                .Worksheets("Site SDC Data").Range("soil").Value = CType(site_soil_d, String)
                .Worksheets("Site SDC Data").Range("risk").Value = CType(risk_category_d, String)
                .Worksheets("Site SDC Data").Range("ss").ClearContents
                .Worksheets("Site SDC Data").Range("suno").ClearContents
                .Worksheets("Site SDC Data").Range("tl").ClearContents
            Else
                If Not IsNothing(Me.use_asce) Then
                    .Worksheets("Reference").Range("Use_ASCE").Value = CType(Me.use_asce, Boolean)
                End If
                If Not IsNothing(Me.site_soil) Then
                    .Worksheets("Site SDC Data").Range("soil").Value = CType(Me.site_soil, String)
                End If
                If Not IsNothing(Me.risk_category) Then
                    .Worksheets("Site SDC Data").Range("risk").Value = CType(Me.risk_category, String)
                End If
                If Not IsNothing(Me.ss) Then
                    .Worksheets("Site SDC Data").Range("ss").Value = CType(Me.ss, Double)
                Else
                    .Worksheets("Site SDC Data").Range("ss").ClearContents
                End If
                If Not IsNothing(Me.s1) Then
                    .Worksheets("Site SDC Data").Range("suno").Value = CType(Me.s1, Double)
                Else
                    .Worksheets("Site SDC Data").Range("suno").ClearContents
                End If
                If Not IsNothing(Me.tl) Then
                    .Worksheets("Site SDC Data").Range("tl").Value = CType(Me.tl, Double)
                Else
                    .Worksheets("Site SDC Data").Range("tl").ClearContents
                End If
            End If

            If Not IsNothing(Me.importance_factor_override) Then
                .Worksheets("Reference").Range("ie_override").Value = CType(Me.importance_factor_override, Boolean)
            End If
            If Not IsNothing(Me.importance_factor_user) Then
                .Worksheets("Site SDC Data").Range("ie_user").Value = CType(Me.importance_factor_user, Double)
            Else
                .Worksheets("Site SDC Data").Range("ie_user").ClearContents
            End If
            If Not IsNothing(Me.response_accel_override) Then
                .Worksheets("Reference").Range("sd_override").Value = CType(Me.response_accel_override, Boolean)
            End If
            If Not IsNothing(Me.sds_user) Then
                .Worksheets("Site SDC Data").Range("sds_user").Value = CType(Me.sds_user, Double)
            Else
                .Worksheets("Site SDC Data").Range("sds_user").ClearContents
            End If
            If Not IsNothing(Me.sd1_user) Then
                .Worksheets("Site SDC Data").Range("sduno_user").Value = CType(Me.sd1_user, Double)
            Else
                .Worksheets("Site SDC Data").Range("sduno_user").ClearContents
            End If
            If Not IsNothing(Me.amp_factor) Then
                .Worksheets("Tower Data").Range("As").Value = CType(Me.amp_factor, Double)
            Else
                .Worksheets("Tower Data").Range("As").ClearContents
            End If
            If Not IsNothing(Me.tia_approx_period) Then
                .Worksheets("Reference").Range("Period").Value = CType(Me.tia_approx_period, Boolean)
            End If
            If Not IsNothing(Me.fundamental_period_user) Then
                .Worksheets("Tower Data").Range("t_user").Value = CType(Me.fundamental_period_user, Double)
            Else
                .Worksheets("Tower Data").Range("t_user").ClearContents
            End If
            If Not IsNothing(Me.mp_density_override) Then
                .Worksheets("Tower Data").Range("mp_d_override").Value = CType(Me.mp_density_override, Boolean)
            End If
            If Not IsNothing(Me.density_tower_material) Then
                .Worksheets("Tower Data").Range("D_mp").Value = CType(Me.density_tower_material, Double)
            Else
                .Worksheets("Tower Data").Range("D_mp").ClearContents
            End If
            If Not IsNothing(Me.elasticity) Then
                .Worksheets("Tower Data").Range("E").Value = CType(Me.elasticity, Double)
            Else
                .Worksheets("Tower Data").Range("E").ClearContents
            End If
            If Not IsNothing(Me.create_seismic_loads) Then
                .Worksheets("Reference").Range("Existing_UF").Value = CType(Me.create_seismic_loads, Boolean)
            End If
            If Not IsNothing(Me.user_force_appurtenance) Then
                .Worksheets("Reference").Range("Existing_UF_Apps").Value = CType(Me.user_force_appurtenance, Boolean)
            End If

        End With

    End Sub

#End Region

#Region "Save to EDS"

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Version.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lat_sign.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lat_deg.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lat_min.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lat_sec.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.long_sign.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.long_deg.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.long_min.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.long_sec.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.use_asce.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.site_soil.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.risk_category.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ss.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.s1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tl.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.importance_factor_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.importance_factor_user.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.response_accel_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.sds_user.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.sd1_user.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.amp_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tia_approx_period.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fundamental_period_user.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.mp_density_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.density_tower_material.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.elasticity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.create_seismic_loads.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.user_force_appurtenance.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.design_code.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        SQLInsertFields = SQLInsertFields.AddtoDBString("lat_sign")
        SQLInsertFields = SQLInsertFields.AddtoDBString("lat_deg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("lat_min")
        SQLInsertFields = SQLInsertFields.AddtoDBString("lat_sec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("long_sign")
        SQLInsertFields = SQLInsertFields.AddtoDBString("long_deg")
        SQLInsertFields = SQLInsertFields.AddtoDBString("long_min")
        SQLInsertFields = SQLInsertFields.AddtoDBString("long_sec")
        SQLInsertFields = SQLInsertFields.AddtoDBString("use_asce")
        SQLInsertFields = SQLInsertFields.AddtoDBString("site_soil")
        SQLInsertFields = SQLInsertFields.AddtoDBString("risk_category")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ss")
        SQLInsertFields = SQLInsertFields.AddtoDBString("s1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tl")
        SQLInsertFields = SQLInsertFields.AddtoDBString("importance_factor_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("importance_factor_user")
        SQLInsertFields = SQLInsertFields.AddtoDBString("response_accel_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sds_user")
        SQLInsertFields = SQLInsertFields.AddtoDBString("sd1_user")
        SQLInsertFields = SQLInsertFields.AddtoDBString("amp_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tia_approx_period")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fundamental_period_user")
        SQLInsertFields = SQLInsertFields.AddtoDBString("mp_density_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("density_tower_material")
        SQLInsertFields = SQLInsertFields.AddtoDBString("elasticity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("create_seismic_loads")
        SQLInsertFields = SQLInsertFields.AddtoDBString("user_force_appurtenance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("design_code")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tool_version = " & Me.Version.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit = " & Me.bus_unit.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id = " & Me.structure_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lat_sign = " & Me.lat_sign.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lat_deg = " & Me.lat_deg.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lat_min = " & Me.lat_min.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lat_sec = " & Me.lat_sec.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("long_sign = " & Me.long_sign.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("long_deg = " & Me.long_deg.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("long_min = " & Me.long_min.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("long_sec = " & Me.long_sec.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("use_asce = " & Me.use_asce.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("site_soil = " & Me.site_soil.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("risk_category = " & Me.risk_category.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ss = " & Me.ss.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("s1 = " & Me.s1.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tl = " & Me.tl.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("importance_factor_override = " & Me.importance_factor_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("importance_factor_user = " & Me.importance_factor_user.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("response_accel_override = " & Me.response_accel_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sds_user = " & Me.sds_user.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("sd1_user = " & Me.sd1_user.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("amp_factor = " & Me.amp_factor.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tia_approx_period = " & Me.tia_approx_period.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("fundamental_period_user = " & Me.fundamental_period_user.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("mp_density_override = " & Me.mp_density_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("density_tower_material = " & Me.density_tower_material.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("elasticity = " & Me.elasticity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("create_seismic_loads = " & Me.create_seismic_loads.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("user_force_appurtenance = " & Me.user_force_appurtenance.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("design_code = " & Me.design_code.ToString.FormatDBValue)

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
        Dim otherToCompare As CCISeismic = TryCast(other, CCISeismic)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.Version.CheckChange(otherToCompare.Version, changes, categoryName, "Tool Version"), Equals, False)
        'Equals = If(Me.bus_unit.CheckChange(otherToCompare.bus_unit, changes, categoryName, "Bus Unit"), Equals, False)
        'Equals = If(Me.structure_id.CheckChange(otherToCompare.structure_id, changes, categoryName, "Structure Id"), Equals, False)
        'Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        'Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)
        Equals = If(Me.lat_sign.CheckChange(otherToCompare.lat_sign, changes, categoryName, "Lat Sign"), Equals, False)
        Equals = If(Me.lat_deg.CheckChange(otherToCompare.lat_deg, changes, categoryName, "Lat Deg"), Equals, False)
        Equals = If(Me.lat_min.CheckChange(otherToCompare.lat_min, changes, categoryName, "Lat Min"), Equals, False)
        Equals = If(Me.lat_sec.CheckChange(otherToCompare.lat_sec, changes, categoryName, "Lat Sec"), Equals, False)
        Equals = If(Me.long_sign.CheckChange(otherToCompare.long_sign, changes, categoryName, "Long Sign"), Equals, False)
        Equals = If(Me.long_deg.CheckChange(otherToCompare.long_deg, changes, categoryName, "Long Deg"), Equals, False)
        Equals = If(Me.long_min.CheckChange(otherToCompare.long_min, changes, categoryName, "Long Min"), Equals, False)
        Equals = If(Me.long_sec.CheckChange(otherToCompare.long_sec, changes, categoryName, "Long Sec"), Equals, False)
        Equals = If(Me.use_asce.CheckChange(otherToCompare.use_asce, changes, categoryName, "Use Asce"), Equals, False)
        Equals = If(Me.site_soil.CheckChange(otherToCompare.site_soil, changes, categoryName, "Site Soil"), Equals, False)
        Equals = If(Me.risk_category.CheckChange(otherToCompare.risk_category, changes, categoryName, "Risk Category"), Equals, False)
        Equals = If(Me.ss.CheckChange(otherToCompare.ss, changes, categoryName, "Ss"), Equals, False)
        Equals = If(Me.s1.CheckChange(otherToCompare.s1, changes, categoryName, "S1"), Equals, False)
        Equals = If(Me.tl.CheckChange(otherToCompare.tl, changes, categoryName, "Tl"), Equals, False)
        Equals = If(Me.importance_factor_override.CheckChange(otherToCompare.importance_factor_override, changes, categoryName, "Importance Factor Override"), Equals, False)
        Equals = If(Me.importance_factor_user.CheckChange(otherToCompare.importance_factor_user, changes, categoryName, "Importance Factor User"), Equals, False)
        Equals = If(Me.response_accel_override.CheckChange(otherToCompare.response_accel_override, changes, categoryName, "Response Accel Override"), Equals, False)
        Equals = If(Me.sds_user.CheckChange(otherToCompare.sds_user, changes, categoryName, "Sds User"), Equals, False)
        Equals = If(Me.sd1_user.CheckChange(otherToCompare.sd1_user, changes, categoryName, "Sd1 User"), Equals, False)
        Equals = If(Me.amp_factor.CheckChange(otherToCompare.amp_factor, changes, categoryName, "Amp Factor"), Equals, False)
        Equals = If(Me.tia_approx_period.CheckChange(otherToCompare.tia_approx_period, changes, categoryName, "Tia Approx Period"), Equals, False)
        Equals = If(Me.fundamental_period_user.CheckChange(otherToCompare.fundamental_period_user, changes, categoryName, "Fundamental Period User"), Equals, False)
        Equals = If(Me.mp_density_override.CheckChange(otherToCompare.mp_density_override, changes, categoryName, "MP Density Override"), Equals, False)
        Equals = If(Me.density_tower_material.CheckChange(otherToCompare.density_tower_material, changes, categoryName, "Density Tower Material"), Equals, False)
        Equals = If(Me.elasticity.CheckChange(otherToCompare.elasticity, changes, categoryName, "Elasticity"), Equals, False)
        Equals = If(Me.create_seismic_loads.CheckChange(otherToCompare.create_seismic_loads, changes, categoryName, "Create Seismic Loads"), Equals, False)
        Equals = If(Me.user_force_appurtenance.CheckChange(otherToCompare.user_force_appurtenance, changes, categoryName, "User Force Appurtenance"), Equals, False)
        Equals = If(Me.design_code.CheckChange(otherToCompare.design_code, changes, categoryName, "Design Code"), Equals, False)

        Return Equals

    End Function
#End Region

End Class
