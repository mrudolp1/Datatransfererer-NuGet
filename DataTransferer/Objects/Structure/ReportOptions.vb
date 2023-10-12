﻿Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.Serialization

<DataContractAttribute()>
Public Class ReportOptions
    Inherits EDSObjectWithQueries

#Region "Properties"

    'Required overriden Properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Report Options"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "report.report_options"
        End Get
    End Property

    'Properties: Store in SQL report.report_options table (PK = WO)
    <DataMember()> Public Property FromDatabaseWO As String 'Not actual WO; just whatever we get from the Database (so for default options, could be old WO)
    <DataMember()> Public Property ReportType As String
    <DataMember()> Public Property ConfigurationType As String
    <DataMember()> Public Property LC As String
    <DataMember()> Public Property LCSubtype As String
    <DataMember()> Public Property CodeRef As String
    <DataMember()> Public Property ToBeStamped As Boolean
    <DataMember()> Public Property ToBeGivenToCustomer As Boolean
    <DataMember()> Public Property OnlySuperStructureAnalyzed As Boolean
    <DataMember()> Public Property NewBuildInNewCode As Boolean
    <DataMember()> Public Property IBM As Boolean
    <DataMember()> Public Property TacExposureChange As Boolean
    <DataMember()> Public Property TacTopoChange As Boolean
    <DataMember()> Public Property MappingDocuments As Boolean
    <DataMember()> Public Property ProposedExtension As Boolean
    <DataMember()> Public Property ExtensionHeight As Double
    <DataMember()> Public Property CanisterExtension As Boolean
    <DataMember()> Public Property IsModified As Boolean
    <DataMember()> Public Property RohnPirodFlangePlates As Boolean
    <DataMember()> Public Property FlangeFEA As Boolean
    <DataMember()> Public Property MpSliceOption As Integer '0,1,2
    <DataMember()> Public Property ConditionallyPassing As Boolean
    <DataMember()> Public Property GradeBeamAnalysisNeeded As Boolean
    <DataMember()> Public Property GradeBeamsRequired As Boolean
    <DataMember()> Public Property GroutRequired As Boolean
    'Workflow bug 189 
    'Added RCTInlucded option 
    <DataMember()> Public Property RCTIncluded As Boolean
    'Workflow bug 10
    'added Concealment_Tower option
    <DataMember()> Public Property ConcealmentTower As Boolean
    'Workflow bug 138
    'added mapped_by 
    <DataMember()> Public Property MappedBy As String
    'Workflow bug 204
    'Added statewording option
    <DataMember()> Public Property UseStateWording As Boolean

    <DataMember()> Public Property ATTAddendum As Boolean
    <DataMember()> Public Property RemoveCFDAreas As Boolean = True
    <DataMember()> Public Property UseTiltTwistWording As Boolean
    <DataMember()> Public Property LicenseOnly As Boolean
    <DataMember()> Public Property MultipleFoundationsConsidered As Boolean
    <DataMember()> Public Property RohnClips As Boolean
    <DataMember()> Public Property TopographicCategoryOtherThan1 As Boolean
    <DataMember()> Public Property ImportanceFactorOtherThan1 As Boolean
    <DataMember()> Public Property PrevWO As Integer
    <DataMember()> Public Property IsDefault As Boolean

    'Properties: Get from EDS based on WO(???)
    <DataMember()> Public Property EngName As String
    <DataMember()> Public Property EngQAName As String
    <DataMember()> Public Property EngStampName As String
    <DataMember()> Public Property EngStampTitle As String

    <DataMember()> Public Property ReportDate As Date = Today
    <DataMember()> Public Property JurisdictionWording As String
    <DataMember()> Public Property IsEditMode As Boolean


    'Lists: Stored in db under report.report_lists
    <DataMember()> Public Property Assumptions As BindingList(Of String) = New BindingList(Of String)
    'From
    '{"Tower and structures were maintained in accordance with the TIA-222 Standard.", "The configuration of antennas, transmission cables, mounts and other appurtenances are as specified in Tables 1 and 2 and the referenced drawings."}
    <DataMember()> Public Property Notes As BindingList(Of String) = New BindingList(Of String)
    <DataMember()> Public Property LoadingChanges As BindingList(Of String) = New BindingList(Of String)

    'Equipment (Tables 1,2,3)
    <DataMember()> Public Property ProposedEquipment As List(Of Equipment) = New List(Of Equipment) 'Table 1
    <DataMember()> Public Property ConditionalEquipment As List(Of Equipment) = New List(Of Equipment) 'Table 2
    <DataMember()> Public Property OtherEquipment As List(Of Equipment) = New List(Of Equipment)    'Table 3

    'Documents (for Table 4)
    <DataMember()> Public Property TableDocuments As List(Of TableDocument) = New List(Of TableDocument) 'Table 4

    'Capacities (for Table 5)
    <DataMember()> Public Property OtherCapacities As List(Of Capacity) = New List(Of Capacity) 'Table 5

    'Temporary place to put LMP data, to eventually be merged with the Equipment Tables in the Report Class Library...
    <DataMember()> Public Property temp_LMP As List(Of FeedLineInformation) = New List(Of FeedLineInformation)

    'File management
    <DataMember()> Public Property WorkingDir As DirectoryInfo
    <DataMember()> Public Property ReportDir As DirectoryInfo

    'Files / Appendixes
    <DataMember()> Public Property Files As Dictionary(Of String, List(Of FilepathWithPriority)) = New Dictionary(Of String, List(Of FilepathWithPriority))()


    'Helper Variables
    <DataMember()> Public Property IsFromDB As Boolean
    <DataMember()> Public Property IsFromDefault As Boolean

#End Region

#Region "ToString functions"
    Public Function AssumptionsString() As String
        Dim result As String = ""

        For i As Integer = 0 To Assumptions.Count() - 2
            result += Assumptions(i) & vbCrLf
        Next

        If Assumptions.Count() > 0 Then
            result += Assumptions(Assumptions.Count() - 1)
        End If

        Return result

    End Function


    Public Function NotesString() As String
        Dim result As String = ""

        For i As Integer = 0 To Notes.Count() - 2
            result += Notes(i) & vbCrLf
        Next

        If Notes.Count() > 0 Then
            result += Notes(Notes.Count() - 1)
        End If

        Return result

    End Function

    Public Function LoadingChangesString() As String
        Dim result As String = ""

        For i As Integer = 0 To LoadingChanges.Count() - 2
            result += LoadingChanges(i) & vbCrLf
        Next

        If LoadingChanges.Count() > 0 Then
            result += LoadingChanges(LoadingChanges.Count() - 1)
        End If

        Return result

    End Function

    Public Function StatusString() As String
        If IsFromDB Then
            If IsFromDefault Then
                Return "Default options associated with BU " + bus_unit + " and SID " + structure_id + " found." + Environment.NewLine + "Options were loaded from previous WO " & FromDatabaseWO & "." + Environment.NewLine + "No in-progress report was found."
            Else
                Return "Found in-progress report options with WO " + work_order_seq_num + "." + Environment.NewLine + "In-progress options loaded."
            End If
        Else
            Return "No in-progress report or default options were found." + Environment.NewLine + "Report populated with basic options."
        End If
    End Function

#End Region

#Region "Constructors (loading from EDS)"

    Public Sub New() 'Default

    End Sub

    Public Sub New(WorkingDir As String, ReportDir As String, ByVal Parent As EDSObject)
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Me.WorkingDir = New DirectoryInfo(WorkingDir)
        Me.ReportDir = New DirectoryInfo(ReportDir)

        Initialize()
    End Sub

    Public Sub New(BU As String, SID As String, WO As String, WorkingDir As String, ReportDir As String)
        bus_unit = BU
        structure_id = SID
        work_order_seq_num = WO

        If WorkingDir = Nothing Then
            Me.WorkingDir = Nothing
        Else
            Me.WorkingDir = New DirectoryInfo(WorkingDir)
        End If

        If ReportDir = Nothing Then
            Me.ReportDir = Nothing
        Else
            Me.ReportDir = New DirectoryInfo(ReportDir)
        End If

        Initialize()

    End Sub

    Public Sub Initialize()
        IsFromDB = False 'Me.ParentStructure.ReportOptions.
        IsFromDefault = False 'Me.ParentStructure.ReportOptions.

        'Try to load in-progress report
        Dim query1 = "SELECT * FROM report.report_options WHERE work_order_seq_num = '" & work_order_seq_num & "'"
        Using strDS As New DataSet
            sqlLoader(query1, strDS, activeDatabase, databaseIdentity, 500)
            If (strDS.Tables(0).Rows.Count > 0) Then
                IsFromDB = True
                'IsFromDefault = False

                Generate(strDS.Tables(0).Rows(0))
                Return
            End If
        End Using

        'Try to load default options
        Dim query2 = "SELECT * FROM report.report_options WHERE bus_unit='" & bus_unit & "' AND structure_id='" & structure_id & "' AND is_default = 1"
        Using strDS As New DataSet
            sqlLoader(query2, strDS, activeDatabase, databaseIdentity, 500)
            If (strDS.Tables(0).Rows.Count > 0) Then
                'IsFromDB = True
                IsFromDefault = True

                Generate(strDS.Tables(0).Rows(0))
                Return
            End If
        End Using

        'If can't load anything from EDS: Clean slate, bring in normal stuff that every report needs.
        IsFromDB = False
        IsFromDefault = False
        LoadDocumentsFromOracle()
        LoadFeedLinesFromOracle()
    End Sub

    'Load everything from EDS data tables (in progress report)
    Public Sub Generate(ByVal SiteCodeDataRow As DataRow)

#Region "Items"
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("work_order_seq_num"), String)) Then
                Me.FromDatabaseWO = CType(SiteCodeDataRow.Item("work_order_seq_num"), String)
            Else
                Me.FromDatabaseWO = Nothing
            End If
        Catch ex As Exception
            Me.FromDatabaseWO = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("bus_unit"), String)) Then
                Me.bus_unit = CType(SiteCodeDataRow.Item("bus_unit"), String)
            Else
                Me.bus_unit = Nothing
            End If
        Catch ex As Exception
            Me.bus_unit = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("structure_id"), String)) Then
                Me.structure_id = CType(SiteCodeDataRow.Item("structure_id"), String)
            Else
                Me.structure_id = Nothing
            End If
        Catch ex As Exception
            Me.structure_id = Nothing
        End Try

        '''3 ITEMS?


        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("report_type"), String)) Then
                Me.ReportType = CType(SiteCodeDataRow.Item("report_type"), String)
            Else
                Me.ReportType = Nothing
            End If
        Catch ex As Exception
            Me.ReportType = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("configuration_type"), String)) Then
                Me.ConfigurationType = CType(SiteCodeDataRow.Item("configuration_type"), String)
            Else
                Me.ConfigurationType = Nothing
            End If
        Catch ex As Exception
            Me.ConfigurationType = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("load_configuration"), String)) Then
                Me.LC = CType(SiteCodeDataRow.Item("load_configuration"), String)
            Else
                Me.LC = Nothing
            End If
        Catch ex As Exception
            Me.LC = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("code_references"), String)) Then
                Me.CodeRef = CType(SiteCodeDataRow.Item("code_references"), String)
            Else
                Me.CodeRef = Nothing
            End If
        Catch ex As Exception
            Me.CodeRef = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("to_be_stamped"), String)) Then
                Me.ToBeStamped = CType(SiteCodeDataRow.Item("to_be_stamped"), String)
            Else
                Me.ToBeStamped = Nothing
            End If
        Catch ex As Exception
            Me.ToBeStamped = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("to_be_given_to_customer"), String)) Then
                Me.ToBeGivenToCustomer = CType(SiteCodeDataRow.Item("to_be_given_to_customer"), String)
            Else
                Me.ToBeGivenToCustomer = Nothing
            End If
        Catch ex As Exception
            Me.ToBeGivenToCustomer = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("only_superstructure_analyzed"), String)) Then
                Me.OnlySuperStructureAnalyzed = CType(SiteCodeDataRow.Item("only_superstructure_analyzed"), String)
            Else
                Me.OnlySuperStructureAnalyzed = Nothing
            End If
        Catch ex As Exception
            Me.OnlySuperStructureAnalyzed = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("new_build_in_new_code"), String)) Then
                Me.NewBuildInNewCode = CType(SiteCodeDataRow.Item("new_build_in_new_code"), String)
            Else
                Me.NewBuildInNewCode = Nothing
            End If
        Catch ex As Exception
            Me.NewBuildInNewCode = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("IBM"), String)) Then
                Me.IBM = CType(SiteCodeDataRow.Item("IBM"), String)
            Else
                Me.IBM = Nothing
            End If
        Catch ex As Exception
            Me.IBM = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("tac_exposure_change"), String)) Then
                Me.TacExposureChange = CType(SiteCodeDataRow.Item("tac_exposure_change"), String)
            Else
                Me.TacExposureChange = Nothing
            End If
        Catch ex As Exception
            Me.TacExposureChange = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("tac_topo_change"), String)) Then
                Me.TacTopoChange = CType(SiteCodeDataRow.Item("tac_topo_change"), String)
            Else
                Me.TacTopoChange = Nothing
            End If
        Catch ex As Exception
            Me.TacTopoChange = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("proposed_tower_extension"), String)) Then
                Me.ProposedExtension = CType(SiteCodeDataRow.Item("proposed_tower_extension"), String)
            Else
                Me.ProposedExtension = Nothing
            End If
        Catch ex As Exception
            Me.ProposedExtension = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("canister_extension"), String)) Then
                Me.CanisterExtension = CType(SiteCodeDataRow.Item("canister_extension"), String)
            Else
                Me.CanisterExtension = Nothing
            End If
        Catch ex As Exception
            Me.CanisterExtension = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("mapping_documents"), String)) Then
                Me.MappingDocuments = CType(SiteCodeDataRow.Item("mapping_documents"), String)
            Else
                Me.MappingDocuments = Nothing
            End If
        Catch ex As Exception
            Me.MappingDocuments = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("is_modified"), String)) Then
                Me.IsModified = CType(SiteCodeDataRow.Item("is_modified"), String)
            Else
                Me.IsModified = Nothing
            End If
        Catch ex As Exception
            Me.IsModified = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("rohn_or_pirod_flange_plates"), String)) Then
                Me.RohnPirodFlangePlates = CType(SiteCodeDataRow.Item("rohn_or_pirod_flange_plates"), String)
            Else
                Me.RohnPirodFlangePlates = Nothing
            End If
        Catch ex As Exception
            Me.RohnPirodFlangePlates = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("flange_fea"), String)) Then
                Me.FlangeFEA = CType(SiteCodeDataRow.Item("flange_fea"), String)
            Else
                Me.FlangeFEA = Nothing
            End If
        Catch ex As Exception
            Me.FlangeFEA = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("mp_slip"), String)) Then
                Me.MpSliceOption = CType(SiteCodeDataRow.Item("mp_slip"), String)
            Else
                Me.MpSliceOption = Nothing
            End If
        Catch ex As Exception
            Me.MpSliceOption = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("conditionally_passing"), String)) Then
                Me.ConditionallyPassing = CType(SiteCodeDataRow.Item("conditionally_passing"), String)
            Else
                Me.ConditionallyPassing = Nothing
            End If
        Catch ex As Exception
            Me.ConditionallyPassing = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("grade_beams_required"), String)) Then
                Me.GradeBeamsRequired = CType(SiteCodeDataRow.Item("grade_beams_required"), String)
            Else
                Me.GradeBeamsRequired = Nothing
            End If
        Catch ex As Exception
            Me.GradeBeamsRequired = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("grout_required"), String)) Then
                Me.GroutRequired = CType(SiteCodeDataRow.Item("grout_required"), String)
            Else
                Me.GroutRequired = Nothing
            End If
        Catch ex As Exception
            Me.GroutRequired = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("rct_included"), String)) Then
                Me.RCTIncluded = CType(SiteCodeDataRow.Item("rct_included"), String)
            Else
                Me.RCTIncluded = Nothing
            End If
        Catch ex As Exception
            Me.RCTIncluded = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("mapped_by"), String)) Then
                Me.MappedBy = CType(SiteCodeDataRow.Item("mapped_by"), String)
            Else
                Me.MappedBy = Nothing
            End If
        Catch ex As Exception
            Me.MappedBy = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("use_state_wording"), String)) Then
                Me.UseStateWording = CType(SiteCodeDataRow.Item("use_state_wording"), String)
            Else
                Me.UseStateWording = Nothing
            End If
        Catch ex As Exception
            Me.UseStateWording = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("concealment_tower"), String)) Then
                Me.ConcealmentTower = CType(SiteCodeDataRow.Item("concealment_tower"), String)
            Else
                Me.ConcealmentTower = Nothing
            End If
        Catch ex As Exception
            Me.ConcealmentTower = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("atat_addendum"), String)) Then
                Me.ATTAddendum = CType(SiteCodeDataRow.Item("atat_addendum"), String)
            Else
                Me.ATTAddendum = Nothing
            End If
        Catch ex As Exception
            Me.ATTAddendum = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("remove_cfd_areas"), String)) Then
                Me.RemoveCFDAreas = CType(SiteCodeDataRow.Item("remove_cfd_areas"), String)
            Else
                Me.RemoveCFDAreas = True
            End If
        Catch ex As Exception
            Me.RemoveCFDAreas = True
        End Try


        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("use_tilt_twist"), String)) Then
                Me.UseTiltTwistWording = CType(SiteCodeDataRow.Item("use_tilt_twist"), String)
            Else
                Me.UseTiltTwistWording = Nothing
            End If
        Catch ex As Exception
            Me.UseTiltTwistWording = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("license_only"), String)) Then
                Me.LicenseOnly = CType(SiteCodeDataRow.Item("license_only"), String)
            Else
                Me.LicenseOnly = Nothing
            End If
        Catch ex As Exception
            Me.LicenseOnly = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("load_configuration_subtype"), String)) Then
                Me.LCSubtype = CType(SiteCodeDataRow.Item("load_configuration_subtype"), String)
            Else
                Me.LCSubtype = Nothing
            End If
        Catch ex As Exception
            Me.LCSubtype = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("extension_height"), String)) Then
                Me.ExtensionHeight = CType(SiteCodeDataRow.Item("extension_height"), String)
            Else
                Me.ExtensionHeight = Nothing
            End If
        Catch ex As Exception
            Me.ExtensionHeight = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("grade_beam_analysis_needed"), String)) Then
                Me.GradeBeamAnalysisNeeded = CType(SiteCodeDataRow.Item("grade_beam_analysis_needed"), String)
            Else
                Me.GradeBeamAnalysisNeeded = Nothing
            End If
        Catch ex As Exception
            Me.GradeBeamAnalysisNeeded = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("multiple_foundations_considered"), String)) Then
                Me.MultipleFoundationsConsidered = CType(SiteCodeDataRow.Item("multiple_foundations_considered"), String)
            Else
                Me.MultipleFoundationsConsidered = Nothing
            End If
        Catch ex As Exception
            Me.MultipleFoundationsConsidered = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("rohn_clips"), String)) Then
                Me.RohnClips = CType(SiteCodeDataRow.Item("rohn_clips"), String)
            Else
                Me.RohnClips = Nothing
            End If
        Catch ex As Exception
            Me.RohnClips = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("topographic_category_other_than_1"), String)) Then
                Me.TopographicCategoryOtherThan1 = CType(SiteCodeDataRow.Item("topographic_category_other_than_1"), String)
            Else
                Me.TopographicCategoryOtherThan1 = Nothing
            End If
        Catch ex As Exception
            Me.TopographicCategoryOtherThan1 = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("prev_wo"), String)) Then
                Me.PrevWO = CType(SiteCodeDataRow.Item("prev_wo"), String)
            Else
                Me.PrevWO = Nothing
            End If
        Catch ex As Exception
            Me.PrevWO = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("report_date"), String)) Then
                '    Me.ReportDate = CType(SiteCodeDataRow.Item("report_date"), String)
                'Else
                Me.ReportDate = Today
            End If
        Catch ex As Exception
            Me.ReportDate = Today
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("custom_jurisdiction_wording"), String)) Then
                Me.JurisdictionWording = CType(SiteCodeDataRow.Item("custom_jurisdiction_wording"), String)
            Else
                Me.JurisdictionWording = Nothing
            End If
        Catch ex As Exception
            Me.JurisdictionWording = Nothing
        End Try

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("is_edit_mode"), String)) Then
                Me.IsEditMode = CType(SiteCodeDataRow.Item("is_edit_mode"), String)
            Else
                Me.IsEditMode = Nothing
            End If
        Catch ex As Exception
            Me.IsEditMode = Nothing
        End Try

        'If Not Me.IsFromDefault Then
        '    Try
        '        If Not IsDBNull(CType(SiteCodeDataRow.Item("root_dir"), String)) Then
        '            Me.RootDir = New DirectoryInfo(CType(SiteCodeDataRow.Item("root_dir"), String))
        '        Else
        '            Me.RootDir = Nothing
        '        End If
        '    Catch ex As Exception
        '        Me.RootDir = Nothing
        '    End Try
        'End If

        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("EngName"), String)) Then
                Me.EngName = CType(SiteCodeDataRow.Item("EngName"), String)
            Else
                Me.EngName = Nothing
            End If
        Catch ex As Exception
            Me.EngName = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("EngQAName"), String)) Then
                Me.EngQAName = CType(SiteCodeDataRow.Item("EngQAName"), String)
            Else
                Me.EngQAName = Nothing
            End If
        Catch ex As Exception
            Me.EngQAName = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("EngStampName"), String)) Then
                Me.EngStampName = CType(SiteCodeDataRow.Item("EngStampName"), String)
            Else
                Me.EngStampName = Nothing
            End If
        Catch ex As Exception
            Me.EngStampName = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("EngStampTitle"), String)) Then
                Me.EngStampTitle = CType(SiteCodeDataRow.Item("EngStampTitle"), String)
            Else
                Me.EngStampTitle = Nothing
            End If
        Catch ex As Exception
            Me.EngStampTitle = Nothing
        End Try
#End Region

#Region "Load related report lists"
        'Load list items
        Dim query = "SELECT * FROM report.report_lists WHERE work_order_seq_num = '" & work_order_seq_num & "'"
        Using strDS As New DataSet
            sqlLoader(query, strDS, activeDatabase, databaseIdentity, 500)
            If (strDS.Tables(0).Rows.Count > 0) Then
                Assumptions.Clear()
                Notes.Clear()
                LoadingChanges.Clear()
            End If

            For Each item In strDS.Tables(0).Rows
                Dim content As String = CType(item.Item("content"), String)
                If item.Item("list_name") = "assumptions" Then
                    Assumptions.Add(content)
                ElseIf item.Item("list_name") = "notes" Then
                    Notes.Add(content)
                ElseIf item.Item("list_name") = "loading_changes" Then
                    LoadingChanges.Add(content)
                End If
            Next

        End Using
#End Region

#Region "Load related report files"
        If Not IsFromDefault Then

            'Load list items
            If (WorkingDir IsNot Nothing) Then 'Logically: if RootDir isn't set, can't load existing appendix documents. If in UI, will reload from WO folder.
                query = "SELECT * FROM report.report_files WHERE work_order_seq_num = '" & work_order_seq_num & "'"

                Using strDS As New DataSet
                    sqlLoader(query, strDS, activeDatabase, databaseIdentity, 500)
                    If (strDS.Tables(0).Rows.Count > 0) Then
                        Files.Clear()
                        Files.Add("CCIPole", New List(Of FilepathWithPriority))
                        Files.Add("rtf", New List(Of FilepathWithPriority))
                        Files.Add("A", New List(Of FilepathWithPriority))
                        Files.Add("B", New List(Of FilepathWithPriority))
                        Files.Add("C", New List(Of FilepathWithPriority))
                        Files.Add("D", New List(Of FilepathWithPriority))
                        Files.Add("Y", New List(Of FilepathWithPriority))
                        Files.Add("Z", New List(Of FilepathWithPriority))
                        Files.Add("Extra", New List(Of FilepathWithPriority))
                    End If

                    For Each item In strDS.Tables(0).Rows
                        Dim appendix As String = CType(item.Item("appendix_name"), String)
                        If Not Files.ContainsKey(appendix) Then
                            Files.Add(appendix, New List(Of FilepathWithPriority))
                        End If

                        Files(appendix).Add(New FilepathWithPriority(
                    -1,
                    Me.WorkingDir.FullName,
                    item.Item("filename").ToString()
                    ))

                    Next

                End Using
            End If
        End If
#End Region

#Region "Load related report documents (for table 4)"
        'Load doc items
        query = "SELECT * FROM report.report_documents WHERE work_order_seq_num = '" & work_order_seq_num & "'"

        Using strDS As New DataSet
            TableDocuments.Clear()
            sqlLoader(query, strDS, activeDatabase, databaseIdentity, 500)
            If (strDS.Tables(0).Rows.Count > 0) Then
                For Each item In strDS.Tables(0).Rows
                    Dim t As TableDocument = New TableDocument(
                            CType(item.Item("doc_name"), String),
                            CType(item.Item("doc_id"), String),
                            CType(item.Item("source"), String),
                            CType(item.Item("valid"), Boolean))
                    t.Enabled = CType(item.Item("checked"), Boolean)
                    TableDocuments.Add(t)

                Next
                LoadDocumentsFromOracle()
            Else
                LoadDocumentsFromOracle()
            End If

        End Using
#End Region

#Region "Load related extra capacites (for table 5)"
        'Load doc items
        query = "SELECT * FROM report.report_capacities WHERE work_order_seq_num = '" & work_order_seq_num & "'"

        Using strDS As New DataSet
            OtherCapacities.Clear()
            sqlLoader(query, strDS, activeDatabase, databaseIdentity, 500)
            If (strDS.Tables(0).Rows.Count > 0) Then
                For Each item In strDS.Tables(0).Rows
                    Dim t As Capacity = New Capacity(
                            CType(item.Item("note"), String),
                            CType(item.Item("component"), String),
                            CType(item.Item("elevation"), Decimal),
                            CType(item.Item("capacity"), Decimal))
                    OtherCapacities.Add(t)

                Next
            End If

        End Using
#End Region

#Region "Load related report equipment (for table 1-3)"
        'Load list items
        query = "SELECT * FROM report.report_equipment WHERE work_order_seq_num = '" & work_order_seq_num & "'"

        Using strDS As New DataSet
            ProposedEquipment.Clear()
            ConditionalEquipment.Clear()
            OtherEquipment.Clear()
            sqlLoader(query, strDS, activeDatabase, databaseIdentity, 500)

            If (strDS.Tables(0).Rows.Count > 0) Then
                For Each item In strDS.Tables(0).Rows
                    Dim t As Equipment = New Equipment(
                        CType(item.Item("mounting_level"), String),
                        CType(item.Item("center_line_elevation"), String),
                        CType(item.Item("num_antennas"), String),
                        CType(item.Item("antenna_manufacturer"), String),
                        CType(item.Item("antenna_model"), String),
                        CType(item.Item("num_feed_lines"), String),
                        CType(item.Item("feed_line_size"), String)
                    )
                    't.enabled = CType(item.Item("checked"), Boolean)

                    If (CType(item.Item("table_num"), Integer) = 1) Then
                        ProposedEquipment.Add(t)
                    ElseIf (CType(item.Item("table_num"), Integer) = 2) Then
                        ConditionalEquipment.Add(t)
                    ElseIf (CType(item.Item("table_num"), Integer) = 3) Then
                        OtherEquipment.Add(t)
                    End If

                Next

            End If

        End Using
#End Region

    End Sub

#End Region

#Region "Loading Oracle Stuff"
    Private Sub LoadFeedLinesFromOracle()
        Using strDS As New DataSet

            Dim lmp_query = "WITH  
                vars AS 
                    --Dual is an empty dummy table. Use this CTE to pull in variables.
                    --(SELECT 2087011 AS work_order_seq_num FROM DUAL), 
                    --(SELECT 2113811 AS work_order_seq_num FROM DUAL), 
                    --(SELECT 2038745 AS wo FROM DUAL), 
                    (SELECT " + work_order_seq_num + " AS wo FROM DUAL), 

                work_order_seq_num AS (
                    SELECT  wos.work_order_seq_num 
                            ,wos.bus_unit 
                            ,wos.structure_id
                            ,wos.eng_app_id
                            ,wos.crrnt_prjct_rvsn_num
                            ,str.bearing_angle
                            ,str.structure_type
                            ,ord.org_seq_num
                    FROM
                            vars
                            ,aim.work_orders wos
                            ,aim.structure str 
                            ,isite.eng_application ord
                    WHERE 
                        wos.work_order_seq_num = vars.wo
                        AND wos.bus_unit = str.bus_unit (+)
                        AND wos.structure_id = str.structure_id (+)
                        AND wos.eng_app_id = ord.eng_app_id (+)
                    ), 
                eCom As
                    (
                        SELECT 
                            c.mount_level 
                            ,c.position
                            ,c.cmpnt_equipment_catalog_id equipment_catalog_id
                            ,c.bus_unit
                            ,c.structure_id                             
                            ,c.status
                            ,'Installed' config
                            ,c.elvtn_num component_cl
                            ,c.cust_org_seq_num
                            ,NULL eng_app_id
                            ,c.quantity_installed_per_antenna quantity
                        FROM aim.installed_component c
                            ,work_order_seq_num wo
                        WHERE c.bus_unit = wo.bus_unit
                        AND c.structure_id = wo.structure_id
                    ),
                pCom As
                    (
                        SELECT 
                            pc.mount_level 
                            ,pc.position                
                            ,pc.cmpnt_equipment_catalog_id equipment_catalog_id
                            ,pc.bus_unit
                            ,pc.structure_id
                            ,pc.status
                            ,'Proposed' config
                            ,pc.elvtn_num component_cl
                            ,pc.cust_org_seq_num
                            ,pc.eng_app_id
                            ,pc.quantity_installed_per_antenna quantity
                        FROM aim.proposed_component pc
                            ,work_order_seq_num wo
                        WHERE pc.bus_unit = wo.bus_unit
                        AND pc.structure_id = wo.structure_id
                        AND pc.merged_ind = 'N'
                    ),
                lmp AS    
                    (SELECT * FROM eCom UNION ALL (SELECT * FROM pCom)),    
                linear AS
                    (
                        SELECT   
                            mfg.org_name database
                            ,ec.model_number || '(' || es.size_text || ')' USName    
                            ,ec.model_number SIName
                            ,CASE
                                --disable equipment with status other than I, P, or A - this will include Not Installed, MLA, and SLA
                                WHEN NOT (lmp.status = 'P' OR lmp.status = 'I' OR lmp.status = 'A')THEN 'FALSE' 
                                --disable installed config when same customer has a proposed config at same mount level
                                WHEN (lmp.config = 'Installed' 
                                    -- customer has proposed equipment
                                    AND lmp.cust_org_seq_num IN (SELECT DISTINCT cust_org_seq_num FROM lmp WHERE config = 'Proposed') 
                                    -- mount centerline has proposed equipment
                                    AND lmp.mount_level IN (SELECT DISTINCT mount_level FROM lmp WHERE config = 'Proposed')) 
                                    THEN 'FALSE'   
                                --disable all but the newest proposed (max order num) if there are multiple proposed apps from the same customer
                                WHEN lmp.config = 'Proposed' AND lmp.eng_app_id NOT IN (SELECT lmp.eng_app_id 
                                                                                        FROM lmp
                                                                                            ,(SELECT cust_org_seq_num
                                                                                                    ,MAX(eng_app_id) AS maxApp
                                                                                            FROM lmp 
                                                                                            GROUP BY cust_org_seq_num) groupedApps
                                                                                        WHERE lmp.eng_app_id = groupedApps.maxApp) THEN 'FALSE'
                                ELSE 'TRUE'
                                END enabled
                            ,lmp.status
                            ,CASE
                                 --delete statuses other than I, P, or A - this will import them as Unassigned and prevent them from being added to the Table 2 of the report
                                WHEN NOT (lmp.status = 'P' OR lmp.status = 'I' OR lmp.status = 'A') THEN NULL 
                                --unassign all but the newest proposed (max order num) if there are multiple proposed apps from the same customer
                                WHEN lmp.config = 'Proposed' AND lmp.eng_app_id NOT IN (SELECT lmp.eng_app_id 
                                                                                FROM lmp
                                                                                    ,(SELECT cust_org_seq_num
                                                                                            ,MAX(eng_app_id) AS maxApp
                                                                                    FROM lmp 
                                                                                    GROUP BY cust_org_seq_num) groupedApps
                                                                                WHERE lmp.eng_app_id = groupedApps.maxApp) THEN NULL 
                                --change proposed to reserved if it's not for the current WO customer
                                WHEN lmp.status = 'P' AND lmp.cust_org_seq_num <> wo.org_seq_num THEN 'R'
                                ELSE lmp.status
                                END ccicode
                            ,CASE WHEN ecx.equipment_category_lkup_code='RND' THEN 'Ar (Round Structural Component)' ELSE 'Af (Flat Structural Component)' END tnxtype
                            ,lmp.quantity  count  
                            ,cust.org_name customer
                            ,lmp.config
                            ,lmp.eng_app_id order_id   
                            ,lmp.mount_level endHeight
                            ,CAST(es.weight_ounce/16 AS DECIMAL(10,2)) selfWeight
                            ,CAST(es.height_inch AS DECIMAL(10,2)) widthOrDiameter
                            ,es.size_text
                        FROM
                            lmp
                            ,work_order_seq_num wo
                            ,equipment.equipment_catalog            ec
                            ,equipment.equipment_specification      es
                            ,aim.org                                mfg
                            ,aim.org                                cust
                            ,equipment.equipment_category_xref      ecx
                            ,app_common.business_entity             be
                            ,aim.site                               s
                        WHERE
                            lmp.equipment_catalog_id = ec.equipment_catalog_id
                            AND ec.manufacturer_org_id = mfg.org_seq_num
                            AND ec.equipment_category_id = ecx.equipment_category_id
                            AND ec.equipment_catalog_id = es.equipment_catalog_id
                            AND lmp.bus_unit = be.bus_unit
                            AND be.lob_code = 'TWR'
                            AND lmp.bus_unit = s.bus_unit
                            AND ecx.equipment_grouping_lkup_code='FEEDLINE'
                            AND wo.bus_unit = be.bus_unit
                            AND wo.bus_unit = s.bus_unit
                            AND lmp.cust_org_seq_num = cust.org_seq_num (+)
                        ORDER BY
                             lmp.mount_level DESC
                            ,lmp.component_cl DESC
                            ,mfg.org_name ASC     
                    )
            --SELECT *
            --FROM pCom
            SELECT 
                enabled
                --,database
                ,usname 
                --,siname
                --,tnxtype 
                ,status
                ,ccicode
                ,SUM(count) as count
                ,endheight
                --,selfweight
                ,size_text
                --,widthordiameter
                --,MAX(customer)
                --,MAX(config)
                --,order_id
            FROM linear 
            GROUP BY
                enabled
                ,database
                ,usname
                ,siname
                ,tnxtype
                ,ccicode
                ,status
                ,endheight
                ,selfweight
                --,widthordiameter
                ,order_id
                ,size_text
            HAVING enabled = 'TRUE'
            ORDER BY 
                endHeight DESC
                ,MAX(config) DESC
                ,(case ccicode 
                    WHEN 'P' THEN 1
                    WHEN 'R' THEN 2
                    WHEN 'I' THEN 3
                    WHEN 'A' THEN 4
                    ELSE 5
                    END) ASC
                ,(case status 
                    WHEN 'P' THEN 1
                    WHEN 'I' THEN 2
                    WHEN 'A' THEN 3
                    ELSE 4
                    END) ASC
                ,order_id DESC    
                ,database ASC
                ,usname ASC
                "

            Dim result = OracleLoader(lmp_query, "LMP", strDS, 3000, "isit")

            If result Then
                If (strDS.Tables("LMP").Rows.Count > 0) Then
                    For Each item In strDS.Tables("LMP").Rows
                        Dim t As FeedLineInformation = New FeedLineInformation(
                                CType(item.Item("enabled"), Boolean),
                                CType(item.Item("size_text"), String),
                                CType(item.Item("status"), String),
                                CType(item.Item("ccicode"), String),
                                CType(item.Item("count"), String),
                                CType(item.Item("endheight"), String))
                        temp_LMP.Add(t)
                    Next
                End If
            End If
        End Using
        'Dim query25 As String = QueryBuilderFromFile("")
        'LoadDocumentsFromOracle()
    End Sub
    Private Sub LoadDocumentsFromOracle()
        Dim Golden As List(Of String) = New List(Of String)({
            "4-GEOTECHNICAL REPORTS",
            "4-TOWER FOUNDATION DRAWINGS/DESIGN/SPECS",
            "4-TOWER MANUFACTURER DRAWINGS",
            "4-POST-INSTALLATION INSPECTION",
            "4-TOWER REINFORCEMENT DESIGN/DRAWINGS/DATA",
            "4-POST-MODIFICATION INSPECTION"})

        Dim doc_query = "select dtm.doc_type_name doc_name, dim.doc_id doc_id, doc_actvy_status_lkup_code validity
                from gds_objects.document_indx_mv dim, gds_objects.document_type_mv dtm, aim.document_activity t
                where dim.ctry_id = 'US'
                    and dim.doc_type_num = dtm.doc_type_num
                    and dim.otg_app_num = dtm.otg_app_num
                    and dim.doc_id=t.doc_id(+)
                    and dim.bus_unit in ('" & bus_unit & "')
                    and dtm.doc_type_name IN ('4-GEOTECHNICAL REPORTS', '4-TOWER MANUFACTURER DRAWINGS', '4-TOWER FOUNDATION DRAWINGS/DESIGN/SPECS')
UNION
select dtm.doc_type_name doc_name, dim.doc_id doc_id, doc_actvy_status_lkup_code validity
                from gds_objects.document_indx_mv dim, gds_objects.document_type_mv dtm, aim.document_activity t
                where dim.ctry_id = 'US'
                    and dim.doc_type_num = dtm.doc_type_num
                    and dim.otg_app_num = dtm.otg_app_num
                    and dim.doc_id=t.doc_id
                    and dim.bus_unit in ('" & bus_unit & "')
                    and dtm.doc_type_name NOT IN ('4-GEOTECHNICAL REPORTS', '4-TOWER MANUFACTURER DRAWINGS', '4-TOWER FOUNDATION DRAWINGS/DESIGN/SPECS') 
                    and dtm.doc_type_name LIKE '4-%'"
        Dim orList As New List(Of TableDocument)

        Using strDS As New DataSet
            OracleLoader(doc_query, "Documents", strDS, 3000, "isit")

            For Each item In strDS.Tables("Documents").Rows
                Dim t As TableDocument =
                    New TableDocument(item("doc_name"), item("doc_id"), "CCISITES", item("validity") = "VALID")
                If (Golden.Contains(t.Document) And t.Valid) Then
                    t.Enabled = True
                End If
                orList.Add(t)

               Dim found As Boolean = False
                For Each doc As TableDocument In Me.TableDocuments
                    If t.Document = doc.Document AndAlso t.Reference = doc.Reference Then 'If Golden.Equals(doc) Then
                        found = True
                        Exit For
                    End If
                Next

                If Not found Then
                    TableDocuments.Add(t)
                End If

            Next
        End Using
    End Sub

#End Region

#Region "Saving"
    Public Function SaveReportOptionsToEds() As Integer

        Try
            'If new default options, make other options not default
            If IsDefault Then 'Update bit
                Dim x = SQLReplace_Default()
                sqlSender(SQLReplace_Default(), activeDatabase, databaseIdentity, 0.ToString)
            End If

            'Find and update (or insert) report options
            Dim query = "SELECT 1 FROM report.report_options WHERE work_order_seq_num = '" & work_order_seq_num & "'"
            Using strDS As New DataSet
                sqlLoader(query, strDS, activeDatabase, databaseIdentity, 500)
                If strDS.Tables(0).Rows.Count > 0 Then 'Update
                    Dim opt_result = sqlSender(SQLUpdate(), activeDatabase, databaseIdentity, 0.ToString)
                    If (Not opt_result) Then
                        Console.WriteLine(SQLUpdate())
                        Return 500
                    End If

                Else
                    Dim opt_result = sqlSender(SQLInsert(), activeDatabase, databaseIdentity, 0.ToString)
                    If (Not opt_result) Then
                        Console.WriteLine(SQLInsert())
                        Return 500
                    End If
                End If
            End Using

#Region "Deal with report option lists"
            Dim commands = New List(Of SqlCommand)

            'Delete all list items associated with WO
            commands.Add(New SqlCommand("DELETE FROM report.report_lists WHERE work_order_seq_num ='" & work_order_seq_num & "'"))


            Dim queryTemplate = "INSERT INTO report.report_lists (work_order_seq_num, list_name, content) VALUES(" & work_order_seq_num & ",@LIST, @VALUE);"

            For Each Item In Assumptions
                Dim command As SqlCommand = New SqlCommand(queryTemplate)
                command.Parameters.Add("@LIST", SqlDbType.VarChar)
                command.Parameters.Add("@VALUE", SqlDbType.VarChar)

                command.Parameters("@LIST").Value = "assumptions"
                command.Parameters("@VALUE").Value = Item.ToString()
                commands.Add(command)

            Next
            For Each Item In Notes
                Dim command As SqlCommand = New SqlCommand(queryTemplate)
                command.Parameters.Add("@LIST", SqlDbType.VarChar)
                command.Parameters.Add("@VALUE", SqlDbType.VarChar)

                command.Parameters("@LIST").Value = "notes"
                command.Parameters("@VALUE").Value = Item.ToString()

                commands.Add(command)
            Next
            For Each Item In LoadingChanges
                Dim command As SqlCommand = New SqlCommand(queryTemplate)
                command.Parameters.Add("@LIST", SqlDbType.VarChar)
                command.Parameters.Add("@VALUE", SqlDbType.VarChar)

                command.Parameters("@LIST").Value = "loading_changes"
                command.Parameters("@VALUE").Value = Item.ToString()

                commands.Add(command)
            Next

            Dim result = safeSqlTransactionSender(commands, activeDatabase, databaseIdentity, 500)
            If (Not result) Then
                Return 500
            End If

#End Region
#Region "Save report option files"

            commands = New List(Of SqlCommand)
            'Delete all appendixes filepaths associated with WO
            commands.Add(New SqlCommand("DELETE FROM report.report_files WHERE work_order_seq_num ='" & work_order_seq_num & "'"))

            'Add current appendix filepaths

            queryTemplate = "INSERT INTO report.report_files (work_order_seq_num, appendix_name, filename) VALUES(" & work_order_seq_num & ",@NAME, @FILENAME);"
            For Each Item In Files
                For Each Document In Item.Value
                    Dim command As SqlCommand = New SqlCommand(queryTemplate)
                    command.Parameters.Add("@NAME", SqlDbType.VarChar)
                    'command.Parameters.Add("@PATH", SqlDbType.VarChar)
                    'command.Parameters.Add("@PRIORITY", SqlDbType.VarChar)
                    command.Parameters.Add("@FILENAME", SqlDbType.VarChar)

                    command.Parameters("@NAME").Value = Item.Key

                    'command.Parameters("@PATH").Value = Document.filepath
                    'command.Parameters("@PRIORITY").Value = Document.priority.ToString()
                    command.Parameters("@FILENAME").Value = Document.filename

                    commands.Add(command)
                Next

            Next

            result = safeSqlTransactionSender(commands, activeDatabase, databaseIdentity, 500)
            If (Not result) Then
                Return 500
            End If

#End Region
#Region "Deal with report documents (Table 4)"


            commands = New List(Of SqlCommand)

            'Delete all document items associated with WO
            commands.Add(New SqlCommand("DELETE FROM report.report_documents WHERE work_order_seq_num ='" & work_order_seq_num & "'"))

            queryTemplate = "INSERT INTO report.report_documents (work_order_seq_num, doc_name, checked, doc_id, valid, source) VALUES(" & work_order_seq_num & ",@DOC_NAME, @CHECKED, @DOC_ID, @VALID, @SOURCE);"

            For Each Item In TableDocuments
                Dim checkedInt As Integer = 0
                If (Item.Enabled) Then
                    checkedInt = 1
                End If
                Dim command As SqlCommand = New SqlCommand(queryTemplate)
                command.Parameters.Add("@DOC_NAME", SqlDbType.VarChar)
                command.Parameters.Add("@CHECKED", SqlDbType.VarChar)
                command.Parameters.Add("@DOC_ID", SqlDbType.VarChar)
                command.Parameters.Add("@VALID", SqlDbType.VarChar)
                command.Parameters.Add("@SOURCE", SqlDbType.VarChar)

                command.Parameters("@DOC_NAME").Value = Item.Document
                command.Parameters("@CHECKED").Value = checkedInt.ToString()
                command.Parameters("@DOC_ID").Value = Item.Reference
                command.Parameters("@VALID").Value = Item.Valid
                command.Parameters("@SOURCE").Value = Item.Source

                commands.Add(command)
            Next

            result = safeSqlTransactionSender(commands, activeDatabase, databaseIdentity, 500)
            If (Not result) Then
                Return 500
            End If

#End Region
#Region "Save other capacity data (Table 5)"

            commands = New List(Of SqlCommand)
            'Delete all document items associated with WO
            commands.Add(New SqlCommand("DELETE FROM report.report_capacities WHERE work_order_seq_num ='" & work_order_seq_num & "'"))

            queryTemplate = "INSERT INTO report.report_capacities (work_order_seq_num,component,elevation,note,capacity) VALUES(" & work_order_seq_num & ",@PARAM1, @PARAM2, @PARAM3, @PARAM4);"
            For Each Item In OtherCapacities
                Dim command As SqlCommand = New SqlCommand(queryTemplate)
                command.Parameters.Add("@PARAM1", SqlDbType.VarChar)
                command.Parameters.Add("@PARAM2", SqlDbType.VarChar)
                command.Parameters.Add("@PARAM3", SqlDbType.VarChar)
                command.Parameters.Add("@PARAM4", SqlDbType.VarChar)

                command.Parameters("@PARAM1").Value = Item.Component
                command.Parameters("@PARAM2").Value = Item.Elevation.ToString()
                command.Parameters("@PARAM3").Value = Item.Notes
                command.Parameters("@PARAM4").Value = Item.cap.ToString()

                commands.Add(command)
            Next

            result = safeSqlTransactionSender(commands, activeDatabase, databaseIdentity, 500)
            If (Not result) Then
                Return 500
            End If
#End Region
#Region "Save report equipment (Table 1,2,3)"
            'Delete all list items associated with WO
            commands = New List(Of SqlCommand)
            'Delete all document items associated with WO
            commands.Add(New SqlCommand("DELETE FROM report.report_equipment WHERE work_order_seq_num ='" & work_order_seq_num & "'"))

            'Delete all document items associated with WO
            queryTemplate = "INSERT INTO report.report_equipment (work_order_seq_num,mounting_level, center_line_elevation, num_antennas, antenna_manufacturer, antenna_model, num_feed_lines, feed_line_size, table_num) VALUES(" & work_order_seq_num & ",@mounting_level, @center_line_elevation, @num_antennas, @antenna_manufacturer, @antenna_model, @num_feed_lines, @feed_line_size, @table_num);"
            For Each Item In ProposedEquipment
                Dim command As SqlCommand = New SqlCommand(queryTemplate)
                command.Parameters.Add("@mounting_level", SqlDbType.VarChar)
                command.Parameters.Add("@center_line_elevation", SqlDbType.VarChar)
                command.Parameters.Add("@num_antennas", SqlDbType.VarChar)
                command.Parameters.Add("@antenna_manufacturer", SqlDbType.VarChar)
                command.Parameters.Add("@antenna_model", SqlDbType.VarChar)
                command.Parameters.Add("@num_feed_lines", SqlDbType.VarChar)
                command.Parameters.Add("@feed_line_size", SqlDbType.VarChar)
                command.Parameters.Add("@table_num", SqlDbType.VarChar)

                command.Parameters("@mounting_level").Value = Item.mounting_level
                command.Parameters("@center_line_elevation").Value = Item.center_line_elevation
                command.Parameters("@num_antennas").Value = Item.num_antennas
                command.Parameters("@antenna_manufacturer").Value = Item.antenna_manufacturer
                command.Parameters("@antenna_model").Value = Item.antenna_model
                command.Parameters("@num_feed_lines").Value = Item.num_feed_lines
                command.Parameters("@feed_line_size").Value = Item.feed_line_size
                command.Parameters("@table_num").Value = 1

                commands.Add(command)
            Next

            For Each Item In ConditionalEquipment
                Dim command As SqlCommand = New SqlCommand(queryTemplate)
                command.Parameters.Add("@mounting_level", SqlDbType.VarChar)
                command.Parameters.Add("@center_line_elevation", SqlDbType.VarChar)
                command.Parameters.Add("@num_antennas", SqlDbType.VarChar)
                command.Parameters.Add("@antenna_manufacturer", SqlDbType.VarChar)
                command.Parameters.Add("@antenna_model", SqlDbType.VarChar)
                command.Parameters.Add("@num_feed_lines", SqlDbType.VarChar)
                command.Parameters.Add("@feed_line_size", SqlDbType.VarChar)
                command.Parameters.Add("@table_num", SqlDbType.VarChar)


                command.Parameters("@mounting_level").Value = Item.mounting_level
                command.Parameters("@center_line_elevation").Value = Item.center_line_elevation
                command.Parameters("@num_antennas").Value = Item.num_antennas
                command.Parameters("@antenna_manufacturer").Value = Item.antenna_manufacturer
                command.Parameters("@antenna_model").Value = Item.antenna_model
                command.Parameters("@num_feed_lines").Value = Item.num_feed_lines
                command.Parameters("@feed_line_size").Value = Item.feed_line_size
                command.Parameters("@table_num").Value = 2

                commands.Add(command)
            Next

            For Each Item In OtherEquipment
                Dim command As SqlCommand = New SqlCommand(queryTemplate)
                command.Parameters.Add("@mounting_level", SqlDbType.VarChar)
                command.Parameters.Add("@center_line_elevation", SqlDbType.VarChar)
                command.Parameters.Add("@num_antennas", SqlDbType.VarChar)
                command.Parameters.Add("@antenna_manufacturer", SqlDbType.VarChar)
                command.Parameters.Add("@antenna_model", SqlDbType.VarChar)
                command.Parameters.Add("@num_feed_lines", SqlDbType.VarChar)
                command.Parameters.Add("@feed_line_size", SqlDbType.VarChar)
                command.Parameters.Add("@table_num", SqlDbType.VarChar)

                command.Parameters("@mounting_level").Value = Item.mounting_level
                command.Parameters("@center_line_elevation").Value = Item.center_line_elevation
                command.Parameters("@num_antennas").Value = Item.num_antennas
                command.Parameters("@antenna_manufacturer").Value = Item.antenna_manufacturer
                command.Parameters("@antenna_model").Value = Item.antenna_model
                command.Parameters("@num_feed_lines").Value = Item.num_feed_lines
                command.Parameters("@feed_line_size").Value = Item.feed_line_size
                command.Parameters("@table_num").Value = 3

                commands.Add(command)
            Next

            result = safeSqlTransactionSender(commands, activeDatabase, databaseIdentity, 500)
            If (Not result) Then
                Return 500
            End If
#End Region

            Return 0
        Catch ex As Exception
            Console.WriteLine(ex)
            Return 500
        End Try

    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = ""
        SQLDelete += "DELETE FROM report.report_documents WHERE work_order_seq_num ='" & work_order_seq_num & "'" + vbCrLf
        SQLDelete += "DELETE FROM report.report_equipment WHERE work_order_seq_num ='" & work_order_seq_num & "'" + vbCrLf
        SQLDelete += "DELETE FROM report.report_capacities WHERE work_order_seq_num ='" & work_order_seq_num & "'" + vbCrLf
        SQLDelete += "DELETE FROM report.report_files WHERE work_order_seq_num ='" & work_order_seq_num & "'" + vbCrLf
        SQLDelete += "DELETE FROM report.report_lists WHERE work_order_seq_num ='" & work_order_seq_num & "'" + vbCrLf
        SQLDelete += "DELETE FROM report.report_options WHERE work_order_seq_num ='" & work_order_seq_num & "' AND is_default <> 'True'" + vbCrLf
        Return SQLDelete
    End Function

    Public Overrides Function SQLInsert() As String
        SQLInsert = "BEGIN" & vbCrLf &
                     "  INSERT INTO [TABLE] ([FIELDS])" & vbCrLf &
                     "  VALUES([VALUES])" & vbCrLf &
                     "END" & vbCrLf
        SQLInsert = SQLInsert.Replace("[TABLE]", Me.EDSTableName)
        SQLInsert = SQLInsert.Replace("[FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.Replace("[VALUES]", Me.SQLInsertValues)
        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = "BEGIN" & vbCrLf &
                  "  Update [TABLE]" &
                  "  SET [UPDATE]" & vbCrLf &
                  "  WHERE work_order_seq_num = [ID]" & vbCrLf &
                  "  [RESULTS]" & vbCrLf &
                  "END" & vbCrLf
        SQLUpdate = SQLUpdate.Replace("[TABLE]", Me.EDSTableName)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.Replace("[ID]", work_order_seq_num)

        SQLUpdate = SQLUpdate.Replace("[RESULTS]", Me.Results.EDSResultQuery)
        Return SQLUpdate
    End Function

    Public Function SQLReplace_Default() As String
        SQLReplace_Default = "BEGIN" & vbCrLf &
                  "  Update [TABLE]" &
                  "  SET [UPDATE]" & vbCrLf &
                  "  WHERE bus_unit = [BU] AND structure_id=[SID] AND is_default = 1" & vbCrLf &
                  "END" & vbCrLf

        SQLReplace_Default = SQLReplace_Default.Replace("[TABLE]", Me.EDSTableName)
        SQLReplace_Default = SQLReplace_Default.Replace("[UPDATE]", "is_default='False'")
        SQLReplace_Default = SQLReplace_Default.Replace("[BU]", bus_unit.NullableToString.FormatDBValue)
        SQLReplace_Default = SQLReplace_Default.Replace("[SID]", structure_id.NullableToString.FormatDBValue)

        Return SQLReplace_Default
    End Function

    'report.reportOptions insert values (for saving new report) (to be used with SQLInsertFields)
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.NullableToString.Replace("'", "''").FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.NullableToString.Replace("'", "''").FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.NullableToString.Replace("'", "''").FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.IsDefault.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").NullableToString.FormatDBValue)

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ATTAddendum.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CanisterExtension.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.CodeRef.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ConditionallyPassing.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ConfigurationType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ExtensionHeight.NullableToString.FormatDBValue)

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.FlangeFEA.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GradeBeamAnalysisNeeded.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GradeBeamsRequired.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.GroutRequired.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.RCTIncluded.NullableToString.FormatDBValue)

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MappedBy.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UseStateWording.NullableToString.FormatDBValue)

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ConcealmentTower.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.IBM.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ImportanceFactorOtherThan1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.LicenseOnly.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.LC.NullableToString.Replace("'", "''").FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.LCSubtype.NullableToString.Replace("'", "''").FormatDBValue)

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MappingDocuments.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MpSliceOption.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MultipleFoundationsConsidered.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.NewBuildInNewCode.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.IsModified.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.OnlySuperStructureAnalyzed.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.PrevWO.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ProposedExtension.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.RemoveCFDAreas.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ReportType.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.RohnClips.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.RohnPirodFlangePlates.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TacExposureChange.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TacTopoChange.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ToBeGivenToCustomer.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ToBeStamped.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.TopographicCategoryOtherThan1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.UseTiltTwistWording.NullableToString.FormatDBValue)

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ReportDate.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.JurisdictionWording.NullableToString.Replace("'", "''").FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.IsEditMode.NullableToString.FormatDBValue)

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.WorkingDir.NullableToString.Replace("'", "''").FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ReportDir.NullableToString.Replace("'", "''").FormatDBValue)

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.EngName.NullableToString.Replace("'", "''").FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.EngQAName.NullableToString.Replace("'", "''").FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.EngStampName.NullableToString.Replace("'", "''").FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.EngStampTitle.NullableToString.Replace("'", "''").FormatDBValue)

        Return SQLInsertValues
    End Function

    'report.reportOptions insert fields (for saving new report)  (to be used with SQLInsertValues)
    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("is_default")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("addDate")

        SQLInsertFields = SQLInsertFields.AddtoDBString("atat_addendum")
        SQLInsertFields = SQLInsertFields.AddtoDBString("canister_extension")
        SQLInsertFields = SQLInsertFields.AddtoDBString("code_references")
        SQLInsertFields = SQLInsertFields.AddtoDBString("conditionally_passing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("configuration_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("extension_height")
        SQLInsertFields = SQLInsertFields.AddtoDBString("flange_fea")
        SQLInsertFields = SQLInsertFields.AddtoDBString("grade_beam_analysis_needed")
        SQLInsertFields = SQLInsertFields.AddtoDBString("grade_beams_required")
        SQLInsertFields = SQLInsertFields.AddtoDBString("grout_required")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rct_included")

        SQLInsertFields = SQLInsertFields.AddtoDBString("mapped_by")
        SQLInsertFields = SQLInsertFields.AddtoDBString("use_state_wording")

        SQLInsertFields = SQLInsertFields.AddtoDBString("concealment_tower")
        SQLInsertFields = SQLInsertFields.AddtoDBString("IBM")
        SQLInsertFields = SQLInsertFields.AddtoDBString("importance_factor_other_than_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("license_only")
        SQLInsertFields = SQLInsertFields.AddtoDBString("load_configuration")
        SQLInsertFields = SQLInsertFields.AddtoDBString("load_configuration_subtype")
        SQLInsertFields = SQLInsertFields.AddtoDBString("mapping_documents")
        SQLInsertFields = SQLInsertFields.AddtoDBString("mp_slip")
        SQLInsertFields = SQLInsertFields.AddtoDBString("multiple_foundations_considered")
        SQLInsertFields = SQLInsertFields.AddtoDBString("new_build_in_new_code")
        SQLInsertFields = SQLInsertFields.AddtoDBString("is_modified")
        SQLInsertFields = SQLInsertFields.AddtoDBString("only_superstructure_analyzed")
        SQLInsertFields = SQLInsertFields.AddtoDBString("prev_wo")
        SQLInsertFields = SQLInsertFields.AddtoDBString("proposed_tower_extension")
        SQLInsertFields = SQLInsertFields.AddtoDBString("remove_cfd_areas")
        SQLInsertFields = SQLInsertFields.AddtoDBString("report_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rohn_clips")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rohn_or_pirod_flange_plates")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tac_exposure_change")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tac_topo_change")
        SQLInsertFields = SQLInsertFields.AddtoDBString("to_be_given_to_customer")
        SQLInsertFields = SQLInsertFields.AddtoDBString("to_be_stamped")
        SQLInsertFields = SQLInsertFields.AddtoDBString("topographic_category_other_than_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("use_tilt_twist")

        SQLInsertFields = SQLInsertFields.AddtoDBString("report_date")
        SQLInsertFields = SQLInsertFields.AddtoDBString("custom_jurisdiction_wording")
        SQLInsertFields = SQLInsertFields.AddtoDBString("is_edit_mode")

        'SQLInsertFields = SQLInsertFields.AddtoDBString("working_dir")
        SQLInsertFields = SQLInsertFields.AddtoDBString("root_dir") 'report_dir

        SQLInsertFields = SQLInsertFields.AddtoDBString("EngName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EngQAName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EngStampName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EngStampTitle")
        Return SQLInsertFields
    End Function

    'report.reportOptions update field fields (for modifying existing report)
    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'Skip wo
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit=" & Me.bus_unit.NullableToString.Replace("'", "''").FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id=" & Me.structure_id.NullableToString.Replace("'", "''").FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("is_default=" & Me.IsDefault.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id=" & Me.modified_person_id.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("addDate=" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").NullableToString.FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("atat_addendum=" & Me.ATTAddendum.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("canister_extension=" & Me.CanisterExtension.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("code_references=" & Me.CodeRef.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("conditionally_passing=" & Me.ConditionallyPassing.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("configuration_type=" & Me.ConfigurationType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("extension_height=" & Me.ExtensionHeight.NullableToString.FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("flange_fea=" & Me.FlangeFEA.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("grade_beam_analysis_needed=" & Me.GradeBeamAnalysisNeeded.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("grade_beams_required=" & Me.GradeBeamsRequired.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("grout_required=" & Me.GroutRequired.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rct_included=" & Me.RCTIncluded.NullableToString.FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("mapped_by=" & Me.MappedBy.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("use_state_wording=" & Me.UseStateWording.NullableToString.FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("concealment_tower=" & Me.ConcealmentTower.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ibm=" & Me.IBM.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("importance_factor_other_than_1=" & Me.ImportanceFactorOtherThan1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("license_only=" & Me.LicenseOnly.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("load_configuration=" & Me.LC.NullableToString.Replace("'", "''").FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("load_configuration_subtype=" & Me.LCSubtype.NullableToString.Replace("'", "''").FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("mapping_documents=" & Me.MappingDocuments.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("mp_slip=" & Me.MpSliceOption.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("multiple_foundations_considered=" & Me.MultipleFoundationsConsidered.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("new_build_in_new_code=" & Me.NewBuildInNewCode.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("is_modified=" & Me.IsModified.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("only_superstructure_analyzed=" & Me.OnlySuperStructureAnalyzed.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("prev_wo=" & Me.PrevWO.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("proposed_tower_extension=" & Me.ProposedExtension.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("remove_cfd_areas=" & Me.RemoveCFDAreas.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("report_type=" & Me.ReportType.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rohn_clips=" & Me.RohnClips.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rohn_or_pirod_flange_plates=" & Me.RohnPirodFlangePlates.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tac_exposure_change=" & Me.TacExposureChange.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tac_topo_change=" & Me.TacTopoChange.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("to_be_given_to_customer=" & Me.ToBeGivenToCustomer.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("to_be_stamped=" & Me.ToBeStamped.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("topographic_category_other_than_1=" & Me.TopographicCategoryOtherThan1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("use_tilt_twist=" & Me.UseTiltTwistWording.NullableToString.FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("report_date=" & Me.ReportDate.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("custom_jurisdiction_wording=" & Me.JurisdictionWording.NullableToString.Replace("'", "''").FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("is_edit_mode=" & Me.IsEditMode.NullableToString.FormatDBValue)

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("working_dir=" & Me.WorkingDir.NullableToString.Replace("'", "''").FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("root_dir=" & Me.ReportDir.NullableToString.Replace("'", "''").FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EngName=" & Me.EngName.NullableToString.Replace("'", "''").FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EngQAName=" & Me.EngQAName.NullableToString.Replace("'", "''").FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EngStampName =" & Me.EngStampName.NullableToString.Replace("'", "''").FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EngStampTitle=" & Me.EngStampTitle.NullableToString.Replace("'", "''").FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Throw New NotImplementedException
    End Function

End Class

#Region "Helper Classes"
<DataContract()>
Public Class FilepathWithPriority
    <DataMember()> Public Property priority As Integer
    <DataMember()> Public Property rootDir As String
    <DataMember()> Public Property filename As String
    <DataMember()> Public Property enabled As Boolean = True

    Public Sub New(ByVal priority As Integer, ByVal rootDir As String, ByVal filename As String)
        Me.priority = priority
        Me.rootDir = rootDir
        Me.filename = filename
    End Sub

    Public Sub New(ByVal root As String, ByVal filename As String)
        Me.priority = -1
        Me.rootDir = root
        Me.filename = filename
        Me.enabled = True
    End Sub

    Public ReadOnly Property filepath As String
        Get
            Return rootDir & "\" & filename
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return filename
    End Function
End Class

<DataContractAttribute()>
Public Class TableDocument
    <DataMember()> Public Property Enabled As Boolean

    <DataMember()> Public Property Document As String
    <DataMember()> Public Property Reference As String
    <DataMember()> Public Property Source As String

    Private _valid As Boolean

    Public ReadOnly Property Valid
        Get
            Return _valid
        End Get
    End Property

    Public Sub New()
        Document = ""
        Reference = ""
        Source = ""
        _valid = True
        Enabled = True
    End Sub


    Public Sub New(doc As String, ref As String, src As String, val As Boolean)
        Document = doc
        Reference = ref
        Source = src
        Try
            _valid = val
        Catch ex As Exception
            _valid = False
        End Try


    End Sub

    Public Function Equals(obj As TableDocument) As Boolean

        If Me.Document <> obj.Document Then Return False
        If Me.Reference <> obj.Reference Then Return False
        If Me.Source <> obj.Source Then Return False
        If Me.Valid <> obj.Valid Then Return False

        Return True
    End Function
End Class

<DataContractAttribute()>
Public Class Equipment
    <Category("EDS"), Description(""), DisplayName("Mounting Level")>
    <DataMember()> Public Property mounting_level As String


    <Category("EDS"), Description(""), DisplayName("Center Line Elevation")>
    <DataMember()> Public Property center_line_elevation As String

    <Category("EDS"), Description(""), DisplayName("Number of Antennas")>
    <DataMember()> Public Property num_antennas As Long

    <Category("EDS"), Description(""), DisplayName("Antenna Manufacturer")>
    <DataMember()> Public Property antenna_manufacturer As String

    <Category("EDS"), Description(""), DisplayName("Antenna Model")>
    <DataMember()> Public Property antenna_model As String

    <Category("EDS"), Description(""), DisplayName("Number of Feed Lines")>
    <DataMember()> Public Property num_feed_lines As String 'Long

    <Category("EDS"), Description(""), DisplayName("Feed Line Size")>
    <DataMember()> Public Property feed_line_size As String

    Public Sub New()
        mounting_level = " - "
        center_line_elevation = " -  "
        num_antennas = 0
        antenna_manufacturer = " - "
        antenna_model = " - "
        num_feed_lines = "" 'num_feed_lines = 0 'Needs to be set to Long to be able to pull an actual number if we want to reference the eri data rather than database value which I think it is currently doing - MRR
        feed_line_size = ""
    End Sub
    Public Sub New(equipment As Equipment)
        mounting_level = equipment.mounting_level
        center_line_elevation = equipment.center_line_elevation
        num_antennas = equipment.num_antennas
        antenna_manufacturer = equipment.antenna_manufacturer
        antenna_model = equipment.antenna_model
        num_feed_lines = equipment.num_feed_lines
        feed_line_size = equipment.feed_line_size


    End Sub


    Public Sub New(
            mounting As String,
            center_line As String,
            num_ant As Long,
            antenna_man As String,
            antenna_mod As String,
            num_feed As String, 'num_feed As Long
            feed_size As String
        )

        mounting_level = mounting
        center_line_elevation = center_line
        num_antennas = num_ant
        antenna_manufacturer = antenna_man
        antenna_model = antenna_mod
        num_feed_lines = num_feed
        feed_line_size = feed_size


    End Sub

    Public Function ToArray() As String()
        Return {
            mounting_level,
            center_line_elevation,
            If(num_antennas = 0, " - ", num_antennas),
            If(antenna_manufacturer, " - "),
            If(antenna_model, " - "),
            num_feed_lines,
            feed_line_size
        }
    End Function

    Public Function IsBlank() As Boolean
        If mounting_level = " - " And
        center_line_elevation = " -  " And
        num_antennas = 0 And
        antenna_manufacturer = " - " And
        antenna_model = " - " And
        num_feed_lines = "" And
        feed_line_size = "" Then
            Return True
        Else
            Return False
        End If

    End Function
End Class

<DataContractAttribute()>
Public Class FeedLineInformation

    <DataMember()> Public Property enabled As Boolean
    <DataMember()> Public Property size As String
    <DataMember()> Public Property status As String
    <DataMember()> Public Property ccicode As String
    <DataMember()> Public Property sum_count As String
    <DataMember()> Public Property endheight As String

    Public Sub New()

    End Sub

    Public Sub New(
            param_enabled As Boolean,
            param_usname As String,
            param_status As String,
            param_ccicode As String,
            param_sum_count As String,
            param_endheight As String
        )

        enabled = param_enabled
        size = param_usname
        status = param_status
        ccicode = param_ccicode
        sum_count = param_sum_count
        endheight = param_endheight


    End Sub

End Class

<DataContractAttribute()>
Public Class Capacity
    <DataMember()> Public Property Notes As String
    <DataMember()> Public Property Component As String
    <DataMember()> Public Property Elevation As Decimal
    <DataMember()> <DisplayName("Capacity")> Public Property cap As Decimal

    Public Sub New()
        Notes = ""
        Component = ""
        Elevation = 0
        Me.cap = 0
    End Sub

    Public Sub New(note As String, component As String, elev As Decimal, cap As Decimal)
        Notes = note
        Me.Component = component
        Elevation = elev
        Me.cap = cap
    End Sub

    Public Function IsBlank() As Boolean
        If Notes = "" And
            Component = "" And
            Elevation = 0 And
            Me.cap = 0 Then
            Return True
        Else
            Return False
        End If

    End Function
End Class

#End Region