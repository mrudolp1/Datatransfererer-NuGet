
'Test cases are created when a test case is selected
'These will correlate to the values in the CSV in the R: drive testing location
Partial Public Class TestCase
    Public Property ID As Integer
    Public Property BU As Integer
    Public Property SID As String
    Public Property WO As Integer
    Public Property COMB As String
    Public Property SAWorkArea As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal csvValue As String())
        Me.ID = csvValue(0)
        Me.BU = csvValue(1)
        Me.SID = csvValue(2)
        Me.WO = csvValue(3)
        Me.COMB = csvValue(4)
        Me.SAWorkArea = csvValue(5)
    End Sub

End Class


Public Class SiteData
    Public Property bus_unit As Integer?
    Public Property structure_id As String
    Public Property work_order_seq_num As Integer?

    Public Sub New()

    End Sub

    Public Sub New(ByVal bu As Integer?, ByVal sid As String, ByVal wo As Integer?)
        Me.bus_unit = bu
        Me.structure_id = sid
        Me.work_order_seq_num = wo
    End Sub

    Public Sub Clear()
        Me.bus_unit = Nothing
        Me.structure_id = Nothing
        Me.work_order_seq_num = Nothing
    End Sub
End Class

Public Enum SyncDirection
    RtoLocal
    LocaltoR
End Enum

Public Enum YesNo
    Yes
    No
End Enum