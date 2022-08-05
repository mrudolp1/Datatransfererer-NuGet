Imports System.ComponentModel

Public Class ReportOptions

    Public Sub New()

    End Sub

   'Properties: Store in SQL report_options table (PK = WO)
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
    Public Property FA As Integer
    
   'Temporary storage: Eventually, store in WO table.
    Public Property customer As String
    Public Property cust_site_num As String
    Public Property cust_site_name As String


   'Properties: Get from EDS (???)
    Public Property EngName As String
    Public Property EngQAName As String
    Public Property EngStampName As String
    Public Property EngStampTitle As String


   'Properties: Do not store in DB (temporary)
    Public Property ReportDate As Date = Today
    Public Property JurisdictionWording As String


   'Lists: Do not store in DB (temporary)
    Public Property Assumptions As BindingList(Of String) = New BindingList(Of String) From
        {"Tower and structures were maintained in accordance with the TIA-222 Standard.", "The configuration of antennas, transmission cables, mounts and other appurtenances are as specified in Tables 1 and 2 and the referenced drawings."}
    Public Property Notes As BindingList(Of String) = New BindingList(Of String)
    Public Property LoadingChanges As BindingList(Of String) = New BindingList(Of String)



    'Helper Functions
    Public Function AssumptionsString() As String
        Dim result As String = ""

        For i As Integer = 0 To Assumptions.Count() - 2
            result += Assumptions(i) & vbCrLf
        Next

        If Assumptions.Count() > 0 
            result += Assumptions(Assumptions.Count()-1)
        End If

        Return result

    End Function

    
    Public Function NotesString() As String
        Dim result As String = ""

        For i As Integer = 0 To Notes.Count() - 2
            result += Notes(i) & vbCrLf
        Next

        If Notes.Count() > 0 
            result += Notes(Notes.Count()-1)
        End If

        Return result

    End Function

    Public Function LoadingChangesString() As String
        Dim result As String = ""

        For i As Integer = 0 To LoadingChanges.Count() - 2
            result += LoadingChanges(i) & vbCrLf
        Next

        If LoadingChanges.Count() > 0 
            result += LoadingChanges(LoadingChanges.Count()-1)
        End If

        Return result

    End Function

End Class