Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Security.Principal

Public Class ReportOptions
    Inherits EDSObjectWithQueries

    'Required overriden Properties
    Public Overrides ReadOnly Property EDSObjectName As String = "Report Options"
    Public Overrides ReadOnly Property EDSTableName As String = "report.report_options"

    Public Property wo As String

    'Properties: Store in SQL report.report_options table (PK = WO)
    Public Property ReportType As String
    Public Property ConfigurationType As String
    Public Property LC As String
    Public Property LCSubtype As String
    Public Property CodeRef As String
    Public Property ToBeStamped As Boolean
    Public Property ToBeGivenToCustomer As Boolean
    Public Property OnlySuperStructureAnalyzed As Boolean
    Public Property NewBuildInNewCode As Boolean
    Public Property IBM As Boolean
    Public Property TacExposureChange As Boolean
    Public Property TacTopoChange As Boolean
    Public Property MappingDocuments As Boolean
    Public Property ProposedExtension As Boolean
    Public Property ExtensionHeight As Double
    Public Property CanisterExtension As Boolean
    Public Property NumModifications As Integer
    Public Property RohnPirodFlangePlates As Boolean
    Public Property FlangeFEA As Boolean
    Public Property MpSliceOption As Integer '0,1,2
    Public Property ConditionallyPassing As Boolean
    Public Property GradeBeamAnalysisNeeded As Boolean
    Public Property GradeBeamsRequired As Boolean
    Public Property GroutRequired As Boolean
    Public Property ATTAddendum As Boolean
    Public Property RemoveCFDAreas As Boolean = True
    Public Property UseTiltTwistWording As Boolean
    Public Property LicenseOnly As Boolean
    Public Property MultipleFoundationsConsidered As Boolean
    Public Property RohnClips As Boolean
    Public Property TopographicCategoryOtherThan1 As Boolean
    Public Property ImportanceFactorOtherThan1 As Boolean
    Public Property PrevWO As Integer


    Public Property IsDefault As Boolean

    'Properties: Get from EDS based on WO(???)
    Public Property EngName As String
    Public Property EngQAName As String
    Public Property EngStampName As String
    Public Property EngStampTitle As String

    Public Property ReportDate As Date = Today
    Public Property JurisdictionWording As String


    'Lists: Stored in db under report.report_lists
    Public Property Assumptions As BindingList(Of String) = New BindingList(Of String) From
        {"Tower and structures were maintained in accordance with the TIA-222 Standard.", "The configuration of antennas, transmission cables, mounts and other appurtenances are as specified in Tables 1 and 2 and the referenced drawings."}
    Public Property Notes As BindingList(Of String) = New BindingList(Of String)
    Public Property LoadingChanges As BindingList(Of String) = New BindingList(Of String)

    'Appendixes
    Public AppendixDocuments As Dictionary(Of String, List(Of FilepathWithPriority)) = New Dictionary(Of String, List(Of FilepathWithPriority))()

    'Helper Variables
    Public Property IsFromDB As Boolean
    Public Property IsFromDefault As Boolean
    Public Property AttemptedWO As String


    'Helper Functions
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
                Return "Default options associated with BU " + bus_unit + " and SID " + structure_id + " were loaded from previous WO " & wo & ".  No in-progress report was found."
            Else
                Return "Found in-progress report options with WO " + AttemptedWO + ".  In-progress options loaded."
            End If
        Else
            Return "No in-progress report or default options were found."
        End If
    End Function

#Region "Constructors"

    Public Sub New() 'Default

    End Sub

    Public Sub New(BU As String, SID As String, WO As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        bus_unit = BU
        structure_id = SID
        AttemptedWO = WO

        Dim query1 = "SELECT * FROM report.report_options WHERE work_order_seq_num = '" & WO & "'"
        Using strDS As New DataSet
            sqlLoader(query1, strDS, ActiveDatabase, LogOnUser, 500)
            If (strDS.Tables(0).Rows.Count > 0) Then
                Generate(strDS.Tables(0).Rows(0), LogOnUser, ActiveDatabase)
                IsFromDB = True
                IsFromDefault = False
                Return
            End If
        End Using

        Dim query2 = "SELECT * FROM report.report_options WHERE bus_unit='" & BU & "' AND structure_id='" & SID & "' AND is_default = 1"
        Using strDS As New DataSet
            sqlLoader(query2, strDS, ActiveDatabase, LogOnUser, 500)
            If (strDS.Tables(0).Rows.Count > 0) Then
                Generate(strDS.Tables(0).Rows(0), LogOnUser, ActiveDatabase)
                IsFromDB = True
                IsFromDefault = True
                Return
            End If
        End Using

        IsFromDB = False

    End Sub

    Public Sub Generate(ByVal SiteCodeDataRow As DataRow, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

#Region "Items"
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("work_order_seq_num"), String)) Then
                Me.wo = CType(SiteCodeDataRow.Item("work_order_seq_num"), String)
            Else
                Me.wo = Nothing
            End If
        Catch ex As Exception
            Me.wo = Nothing
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
            If Not IsDBNull(CType(SiteCodeDataRow.Item("num_modifications"), String)) Then
                Me.NumModifications = CType(SiteCodeDataRow.Item("num_modifications"), String)
            Else
                Me.NumModifications = Nothing
            End If
        Catch ex As Exception
            Me.NumModifications = Nothing
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
                Me.ReportDate = CType(SiteCodeDataRow.Item("report_date"), String)
            Else
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
            If Not IsDBNull(CType(SiteCodeDataRow.Item("EngName"), String)) Then
                Me.EngName= CType(SiteCodeDataRow.Item("EngName"), String)
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
        Dim query = "SELECT * FROM report.report_lists WHERE work_order_seq_num = '" & wo & "'"
        'Console.WriteLine(ActiveDatabase)
        'Console.WriteLine(LogOnUser.AccessToken)
        Using strDS As New DataSet
            sqlLoader(query, strDS, ActiveDatabase, LogOnUser, 500)
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
#Region "Load related report appendixes"
        'Load list items
        query = "SELECT * FROM report.report_appendixes WHERE work_order_seq_num = '" & wo & "'"
        'Console.WriteLine(ActiveDatabase)
        'Console.WriteLine(LogOnUser.AccessToken)

        Using strDS As New DataSet
            sqlLoader(query, strDS, ActiveDatabase, LogOnUser, 500)
            If (strDS.Tables(0).Rows.Count > 0) Then
                AppendixDocuments.Clear()
                AppendixDocuments.Add("rtf", New List(Of FilepathWithPriority))
                AppendixDocuments.Add("A",   New List(Of FilepathWithPriority))
                AppendixDocuments.Add("B",   New List(Of FilepathWithPriority))
                AppendixDocuments.Add("C",   New List(Of FilepathWithPriority))
                AppendixDocuments.Add("D",   New List(Of FilepathWithPriority))
                AppendixDocuments.Add("Y",   New List(Of FilepathWithPriority))
                AppendixDocuments.Add("Z",   New List(Of FilepathWithPriority))
                AppendixDocuments.Add("Extra", New List(Of FilepathWithPriority))
            End If

            For Each item In strDS.Tables(0).Rows
                Dim appendix As String = CType(item.Item("appendix_name"), String)
                If Not AppendixDocuments.ContainsKey(appendix) Then
                    AppendixDocuments.Add(appendix, New List(Of FilepathWithPriority))
                End If

                AppendixDocuments(appendix).Add(New FilepathWithPriority(
                    Integer.Parse(item.Item("priority")),
                    item.Item("filepath").ToString(),
                    item.Item("filename").ToString()
                    ))

            Next

        End Using
#End Region

    End Sub

#End Region

    Public Function SaveReportOptionsToEds(ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String) As Integer
        Try


            'If new default options, make other options not default
            If IsDefault Then 'Update bit
                Dim x = SQLReplace_Default()
                Console.WriteLine(SQLReplace_Default())
                sqlSender(SQLReplace_Default(), ActiveDatabase, LogOnUser, 0.ToString)
            End If

            'Find and update (or insert) report options
            Dim query = "SELECT 1 FROM report.report_options WHERE work_order_seq_num = '" & wo & "'"
            Using strDS As New DataSet
                sqlLoader(query, strDS, ActiveDatabase, LogOnUser, 500)
                If strDS.Tables(0).Rows.Count > 0 Then 'Update
                    sqlSender(SQLUpdate(), ActiveDatabase, LogOnUser, 0.ToString)

                Else 'Insert
                    Console.WriteLine(SQLInsert())
                    sqlSender(SQLInsert(), ActiveDatabase, LogOnUser, 0.ToString)
                    'SQLInsert()
                End If
            End Using

#Region "Deal with report option lists"

            'Delete all list items associated with WO
            query = "DELETE FROM report.report_lists WHERE work_order_seq_num ='" & wo & "'"
            Using strDS As New DataSet
                sqlSender(query, ActiveDatabase, LogOnUser, 500)
            End Using

            'Add current list items
            Using strDS As New DataSet
                Dim FinalQuery = "BEGIN" & vbCrLf &
                                    "[INSERTS]" & vbCrLf &
                                 "END" & vbCrLf
                Dim queryTemplate = "INSERT INTO report.report_lists (work_order_seq_num, list_name, content) VALUES(" & wo & ",'[LIST]', '[VALUE]');"
                Dim inserts = ""
                For Each Item In Assumptions
                    inserts += queryTemplate.Replace("[LIST]", "assumptions").Replace("[VALUE]", Item)
                Next
                For Each Item In Notes
                    inserts += queryTemplate.Replace("[LIST]", "notes").Replace("[VALUE]", Item)
                Next
                For Each Item In LoadingChanges
                    inserts += queryTemplate.Replace("[LIST]", "loading_changes").Replace("[VALUE]", Item)
                Next

                FinalQuery = FinalQuery.Replace("[INSERTS]", inserts)
                Console.WriteLine(FinalQuery)
                sqlSender(FinalQuery, ActiveDatabase, LogOnUser, 500)
            End Using
#End Region
#Region "Deal with report option Appendixes"

            'Delete all appendixes filepaths associated with WO
            query = "DELETE FROM report.report_appendixes WHERE work_order_seq_num ='" & wo & "'"
            Using strDS As New DataSet
                sqlSender(query, ActiveDatabase, LogOnUser, 500)
            End Using

            'Add current appendix filepaths
            Using strDS As New DataSet
                Dim FinalQuery = "BEGIN" & vbCrLf &
                                    "[INSERTS]" & vbCrLf &
                                 "END" & vbCrLf
                Dim queryTemplate = "INSERT INTO report.report_appendixes (work_order_seq_num, appendix_name, filepath, filename, priority) VALUES(" & wo & ",'[NAME]', '[PATH]', '[FILENAME]','[PRIORITY]');"
                Dim inserts = ""
                For Each Item In AppendixDocuments
                    For Each Document In Item.Value
                        inserts += queryTemplate.Replace("[NAME]", Item.Key).Replace("[PATH]", Document.filepath).Replace("[PRIORITY]", Document.priority.ToString()).Replace("[FILENAME]", Document.filename)
                    Next

                Next

                FinalQuery = FinalQuery.Replace("[INSERTS]", inserts)
                Console.WriteLine(FinalQuery)
                sqlSender(FinalQuery, ActiveDatabase, LogOnUser, 500)
            End Using
#End Region


            Return 0
        Catch ex As Exception
            Return 500
        End Try

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
        SQLUpdate = SQLUpdate.Replace("[ID]", wo)

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

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.wo.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.NullableToString.FormatDBValue)
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
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.IBM.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ImportanceFactorOtherThan1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.LicenseOnly.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.LC.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.LCSubtype.NullableToString.FormatDBValue)

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MappingDocuments.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MpSliceOption.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.MultipleFoundationsConsidered.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.NewBuildInNewCode.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.NumModifications.NullableToString.FormatDBValue)
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
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.JurisdictionWording.NullableToString.FormatDBValue)

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.EngName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.EngQAName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.EngStampName.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.EngStampTitle.NullableToString.FormatDBValue)

        Return SQLInsertValues
    End Function

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
        SQLInsertFields = SQLInsertFields.AddtoDBString("IBM")
        SQLInsertFields = SQLInsertFields.AddtoDBString("importance_factor_other_than_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("license_only")
        SQLInsertFields = SQLInsertFields.AddtoDBString("load_configuration")
        SQLInsertFields = SQLInsertFields.AddtoDBString("load_configuration_subtype")
        SQLInsertFields = SQLInsertFields.AddtoDBString("mapping_documents")
        SQLInsertFields = SQLInsertFields.AddtoDBString("mp_slip")
        SQLInsertFields = SQLInsertFields.AddtoDBString("multiple_foundations_considered")
        SQLInsertFields = SQLInsertFields.AddtoDBString("new_build_in_new_code")
        SQLInsertFields = SQLInsertFields.AddtoDBString("num_modifications")
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

        SQLInsertFields = SQLInsertFields.AddtoDBString("EngName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EngQAName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EngStampName")
        SQLInsertFields = SQLInsertFields.AddtoDBString("EngStampTitle")
        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'Skip wo
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit=" & Me.bus_unit.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id=" & Me.structure_id.NullableToString.FormatDBValue)
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
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ibm=" & Me.IBM.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("importance_factor_other_than_1=" & Me.ImportanceFactorOtherThan1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("license_only=" & Me.LicenseOnly.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("load_configuration=" & Me.LC.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("load_configuration_subtype=" & Me.LCSubtype.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("mapping_documents=" & Me.MappingDocuments.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("mp_slip=" & Me.MpSliceOption.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("multiple_foundations_considered=" & Me.MultipleFoundationsConsidered.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("new_build_in_new_code=" & Me.NewBuildInNewCode.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("num_modifications=" & Me.NumModifications.NullableToString.FormatDBValue)
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
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("custom_jurisdiction_wording=" & Me.JurisdictionWording.NullableToString.FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EngName=" & Me.EngName.NullableToString.FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EngQAName=" & Me.EngQAName.NullableToString.FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EngStampName =" & Me.EngStampName.NullableToString.FormatDBValue)

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("EngStampTitle=" & Me.EngStampTitle.NullableToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function


    '
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Throw New NotImplementedException
    End Function

End Class

Public Class FilepathWithPriority
    Public priority As Integer
    Public filepath As String
    Public filename As String
    Public enabled As Boolean = True

    Public Sub New(ByVal priority As Integer, ByVal filepath As String, ByVal filename As String)
        Me.priority = priority
        Me.filepath = filepath
        Me.filename = filename
    End Sub

    Public Sub New(ByVal root As String, ByVal filename As String)
        Me.priority = -1
        Me.filepath = root & "\" & filename
        Me.filename = filename
        Me.enabled = True
    End Sub

    Public Overrides Function ToString() As String
        Return filename
    End Function
End Class
