Imports Microsoft.VisualBasic

Public Class SiteInfo
    Public Sub New()

    End Sub

    'Properties all pulling from CCISites based on bu_num and structure_id
    Public Property bu_num As Integer
    Public Property structure_id As String

    Public Property app_num As String
    Public Property app_rev As String
    Public Property site_name As String
    Public Property jde_num As String
    Public Property site_add As String
    Public Property site_city As String
    Public Property site_county As String
    Public Property site_state As String

    Public Property lat_decimal As Double
    Public ReadOnly Property lat_deg As Integer
        Get
            Return Math.Truncate(lat_decimal)
        End Get
    End Property
    Public ReadOnly Property lat_min As Integer
        Get
            Return Math.Abs(Math.Truncate((lat_decimal - lat_deg) * 60))
        End Get
    End Property
    Public ReadOnly Property lat_sec As Decimal
        Get
            Return Math.Abs(Math.Round(((Math.Abs(lat_decimal - lat_deg)) * 60 - lat_min) * 60, 2))
        End Get
    End Property


    Public Property long_decimal As Double
    Public ReadOnly Property long_deg As Integer
        Get
            Return Math.Truncate(long_decimal)
        End Get
    End Property
    Public ReadOnly Property long_min As Integer
        Get
            Return Math.Abs(Math.Truncate((long_decimal - long_deg) * 60))
        End Get
    End Property
    Public ReadOnly Property long_sec As Decimal
        Get
            Return Math.Abs(Math.Round(((Math.Abs(long_decimal - long_deg)) * 60 - long_min) * 60, 2))
        End Get
    End Property

    Public Property tower_man As String

    'Public Property tower_type As String <-- From TNX
    'Public Property tower_height As String <-- From TNX

End Class
