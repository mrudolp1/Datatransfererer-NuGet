Public Class ReportOptions

    Public Sub New()

    End Sub

    Public Property ReportType As String
    Public Property ConfigurationType As String
    Public Property LC As String
    Public Property CodeRef As String '

    Public Property ToBeStamped As Boolean
    Public Property ToBeGivenToCustomer As Boolean

    Public Property OnlySuperStructureAnalyzed As Boolean
    Public Property NewBuildInNewCode As Boolean

    Public Property IBM As Boolean

    Public Property TacExposureChange As Boolean
    Public Property TacTopoChange As Boolean

    Public Property JurisdictionWording As String


    'Other Property Options
    Public Property MappingDocuments As Boolean
    Public Property CanisterExtension As Boolean
    Public Property ModifiedTower As Boolean
    Public Property RohnPirodFlangePlates As Boolean
    Public Property FlageFEA As Boolean
    Public Property MpSliceLessThanAmount As Boolean
    Public Property ConditionallyPassing As Boolean
    Public Property GradeBeamsRequired As Boolean
    Public Property GroutRequired As Boolean
    Public Property ATATAddendum As Boolean
    Public Property RemoveCFDAreas As Boolean = True

    'UserInfo




End Class