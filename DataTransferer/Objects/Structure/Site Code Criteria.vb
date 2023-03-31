Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports DevExpress.DataAccess.Excel
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient

Partial Public Class SiteCodeCriteria
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Site Code Criteria"

#Region "Define"

    Private _ibc_current As String
    Private _asce_current As String
    Private _tia_current As String
    Private _rev_h_accepted As Boolean?
    Private _rev_h_section_15_5 As Boolean?
    Private _seismic_design_category As String
    Private _frost_depth_tia_g As Double?
    Private _elev_agl As Double?
    Private _topo_category As Integer?
    Private _expo_category As String
    Private _crest_height As Double?
    Private _slope_distance As Double?
    Private _distance_from_crest As Double?
    Private _downwind As Boolean?
    Private _topo_feature As String
    Private _crest_point_elev As Double?
    Private _base_point_elev As Double?
    Private _mid_height_elev As Double?
    Private _crest_to_mid_height_distance As Double?
    Private _tower_point_elev As Double?
    Private _base_kzt As Double?
    Private _site_name As String
    Private _structure_type As String
    Private _eng_app_id As Integer?
    Private _eng_app_id_revision As Integer?
    Private _lat_dec As Double?
    Private _long_dec As Double?

    <Category(""), Description(""), DisplayName("ibc_current")>
    Public Property ibc_current() As String
        Get
            Return Me._ibc_current
        End Get
        Set
            Me._ibc_current = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("asce_current")>
    Public Property asce_current() As String
        Get
            Return Me._asce_current
        End Get
        Set
            Me._asce_current = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("tia_current")>
    Public Property tia_current() As String
        Get
            Return Me._tia_current
        End Get
        Set
            Me._tia_current = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("rev_h_accepted")>
    Public Property rev_h_accepted() As Boolean?
        Get
            Return Me._rev_h_accepted
        End Get
        Set
            Me._rev_h_accepted = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("rev_h_section_15_5")>
    Public Property rev_h_section_15_5() As Boolean?
        Get
            Return Me._rev_h_section_15_5
        End Get
        Set
            Me._rev_h_section_15_5 = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("seismic_design_category")>
    Public Property seismic_design_category() As String
        Get
            Return Me._seismic_design_category
        End Get
        Set
            Me._seismic_design_category = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("frost_depth_tia_g")>
    Public Property frost_depth_tia_g() As Double?
        Get
            Return Me._frost_depth_tia_g
        End Get
        Set
            Me._frost_depth_tia_g = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("elev_agl")>
    Public Property elev_agl() As Double?
        Get
            Return Me._elev_agl
        End Get
        Set
            Me._elev_agl = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("topo_category")>
    Public Property topo_category() As Integer?
        Get
            Return Me._topo_category
        End Get
        Set
            Me._topo_category = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("expo_category")>
    Public Property expo_category() As String
        Get
            Return Me._expo_category
        End Get
        Set
            Me._expo_category = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("crest_height")>
    Public Property crest_height() As Double?
        Get
            Return Me._crest_height
        End Get
        Set
            Me._crest_height = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("slope_distance")>
    Public Property slope_distance() As Double?
        Get
            Return Me._slope_distance
        End Get
        Set
            Me._slope_distance = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("distance_from_crest")>
    Public Property distance_from_crest() As Double?
        Get
            Return Me._distance_from_crest
        End Get
        Set
            Me._distance_from_crest = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("downwind")>
    Public Property downwind() As Boolean?
        Get
            Return Me._downwind
        End Get
        Set
            Me._downwind = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("topo_feature")>
    Public Property topo_feature() As String
        Get
            Return Me._topo_feature
        End Get
        Set
            Me._topo_feature = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("crest_point_elev")>
    Public Property crest_point_elev() As Double?
        Get
            Return Me._crest_point_elev
        End Get
        Set
            Me._crest_point_elev = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("base_point_elev")>
    Public Property base_point_elev() As Double?
        Get
            Return Me._base_point_elev
        End Get
        Set
            Me._base_point_elev = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("mid_height_elev")>
    Public Property mid_height_elev() As Double?
        Get
            Return Me._mid_height_elev
        End Get
        Set
            Me._mid_height_elev = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("crest_to_mid_height_distance")>
    Public Property crest_to_mid_height_distance() As Double?
        Get
            Return Me._crest_to_mid_height_distance
        End Get
        Set
            Me._crest_to_mid_height_distance = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("tower_point_elev")>
    Public Property tower_point_elev() As Double?
        Get
            Return Me._tower_point_elev
        End Get
        Set
            Me._tower_point_elev = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("base_kzt")>
    Public Property base_kzt() As Double?
        Get
            Return Me._base_kzt
        End Get
        Set
            Me._base_kzt = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("site_name")>
    Public Property site_name() As String
        Get
            Return Me._site_name
        End Get
        Set
            Me._site_name = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("structure_type")>
    Public Property structure_type() As String
        Get
            Return Me._structure_type
        End Get
        Set
            Me._structure_type = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("eng_app_id")>
    Public Property eng_app_id() As Integer?
        Get
            Return Me._eng_app_id
        End Get
        Set
            Me._eng_app_id = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("eng_app_id_revision")>
    Public Property eng_app_id_revision() As Integer?
        Get
            Return Me._eng_app_id_revision
        End Get
        Set
            Me._eng_app_id_revision = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("lat_dec")>
    Public Property lat_dec() As Double?
        Get
            Return Me._lat_dec
        End Get
        Set
            Me._lat_dec = Value
        End Set
    End Property
    <Category(""), Description(""), DisplayName("long_dec")>
    Public Property long_dec() As Double?
        Get
            Return Me._long_dec
        End Get
        Set
            Me._long_dec = Value
        End Set
    End Property

#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As SiteCodeCriteria = TryCast(other, SiteCodeCriteria)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.ibc_current.CheckChange(otherToCompare.ibc_current, changes, categoryName, "IBC Current"), Equals, False)
        Equals = If(Me.asce_current.CheckChange(otherToCompare.asce_current, changes, categoryName, "ASCE Current"), Equals, False)
        Equals = If(Me.tia_current.CheckChange(otherToCompare.tia_current, changes, categoryName, "TIA Current"), Equals, False)
        Equals = If(Me.rev_h_accepted.CheckChange(otherToCompare.rev_h_accepted, changes, categoryName, "Rev H Accepted"), Equals, False)
        Equals = If(Me.rev_h_section_15_5.CheckChange(otherToCompare.rev_h_section_15_5, changes, categoryName, "Rev H Section 15 5"), Equals, False)
        Equals = If(Me.seismic_design_category.CheckChange(otherToCompare.seismic_design_category, changes, categoryName, "Seismic Design Category"), Equals, False)
        Equals = If(Me.frost_depth_tia_g.CheckChange(otherToCompare.frost_depth_tia_g, changes, categoryName, "Frost Depth Tia G"), Equals, False)
        Equals = If(Me.elev_agl.CheckChange(otherToCompare.elev_agl, changes, categoryName, "Elev Agl"), Equals, False)
        Equals = If(Me.topo_category.CheckChange(otherToCompare.topo_category, changes, categoryName, "Topo Category"), Equals, False)
        Equals = If(Me.expo_category.CheckChange(otherToCompare.expo_category, changes, categoryName, "Expo Category"), Equals, False)
        Equals = If(Me.crest_height.CheckChange(otherToCompare.crest_height, changes, categoryName, "Crest Height"), Equals, False)
        Equals = If(Me.slope_distance.CheckChange(otherToCompare.slope_distance, changes, categoryName, "Slope Distance"), Equals, False)
        Equals = If(Me.distance_from_crest.CheckChange(otherToCompare.distance_from_crest, changes, categoryName, "Distance From Crest"), Equals, False)
        Equals = If(Me.downwind.CheckChange(otherToCompare.downwind, changes, categoryName, "Downwind"), Equals, False)
        Equals = If(Me.topo_feature.CheckChange(otherToCompare.topo_feature, changes, categoryName, "Topo Feature"), Equals, False)
        Equals = If(Me.crest_point_elev.CheckChange(otherToCompare.crest_point_elev, changes, categoryName, "Crest Point Elev"), Equals, False)
        Equals = If(Me.base_point_elev.CheckChange(otherToCompare.base_point_elev, changes, categoryName, "Base Point Elev"), Equals, False)
        Equals = If(Me.mid_height_elev.CheckChange(otherToCompare.mid_height_elev, changes, categoryName, "Mid Height Elev"), Equals, False)
        Equals = If(Me.crest_to_mid_height_distance.CheckChange(otherToCompare.crest_to_mid_height_distance, changes, categoryName, "Crest To Mid Height Distance"), Equals, False)
        Equals = If(Me.tower_point_elev.CheckChange(otherToCompare.tower_point_elev, changes, categoryName, "Tower Point Elev"), Equals, False)
        Equals = If(Me.base_kzt.CheckChange(otherToCompare.base_kzt, changes, categoryName, "Base Kzt"), Equals, False)
        Equals = If(Me.site_name.CheckChange(otherToCompare.site_name, changes, categoryName, "Site Name"), Equals, False)
        Equals = If(Me.structure_type.CheckChange(otherToCompare.structure_type, changes, categoryName, "Structure Type"), Equals, False)
        Equals = If(Me.eng_app_id.CheckChange(otherToCompare.structure_type, changes, categoryName, "App ID"), Equals, False)
        Equals = If(Me.eng_app_id_revision.CheckChange(otherToCompare.structure_type, changes, categoryName, "App Revision"), Equals, False)
        Equals = If(Me.lat_dec.CheckChange(otherToCompare.lat_dec, changes, categoryName, "Lat Dec"), Equals, False)
        Equals = If(Me.long_dec.CheckChange(otherToCompare.long_dec, changes, categoryName, "Long Dec"), Equals, False)

        Return Equals
    End Function

#Region "Constructors"
    Public Sub New()
        'Variables need to be passed into another constructor
        'Using this just as an example and assuming BU & structure ID exist
    End Sub

    Public Sub New(ByVal SiteCodeDataRow As DataRow)
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
            If Not IsDBNull(CType(SiteCodeDataRow.Item("ibc_current"), String)) Then
                Me.ibc_current = CType(SiteCodeDataRow.Item("ibc_current"), String)
            Else
                Me.ibc_current = Nothing
            End If
        Catch ex As Exception
            Me.ibc_current = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("asce_current"), String)) Then
                Me.asce_current = CType(SiteCodeDataRow.Item("asce_current"), String)
            Else
                Me.asce_current = Nothing
            End If
        Catch ex As Exception
            Me.asce_current = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("tia_current"), String)) Then
                Me.tia_current = CType(SiteCodeDataRow.Item("tia_current"), String)
            Else
                Me.tia_current = Nothing
            End If
        Catch ex As Exception
            Me.tia_current = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("rev_h_accepted"), Boolean)) Then
                Me.rev_h_accepted = CType(SiteCodeDataRow.Item("rev_h_accepted"), Boolean)
            Else
                Me.rev_h_accepted = Nothing
            End If
        Catch ex As Exception
            Me.rev_h_accepted = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("rev_h_section_15_5"), Boolean)) Then
                Me.rev_h_section_15_5 = CType(SiteCodeDataRow.Item("rev_h_section_15_5"), Boolean)
            Else
                Me.rev_h_section_15_5 = Nothing
            End If
        Catch ex As Exception
            Me.rev_h_section_15_5 = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("seismic_design_category"), Boolean)) Then
                Me.seismic_design_category = CType(SiteCodeDataRow.Item("seismic_design_category"), Boolean)
            Else
                Me.seismic_design_category = Nothing
            End If
        Catch ex As Exception
            Me.seismic_design_category = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("frost_depth_tia_g"), Double)) Then
                Me.frost_depth_tia_g = CType(SiteCodeDataRow.Item("frost_depth_tia_g"), Double)
            Else
                Me.frost_depth_tia_g = Nothing
            End If
        Catch ex As Exception
            Me.frost_depth_tia_g = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("elev_agl"), Double)) Then
                Me.elev_agl = CType(SiteCodeDataRow.Item("elev_agl"), Double)
            Else
                Me.elev_agl = Nothing
            End If
        Catch ex As Exception
            Me.elev_agl = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("topo_category"), Integer)) Then
                Me.topo_category = CType(SiteCodeDataRow.Item("topo_category"), Integer)
            Else
                Me.topo_category = Nothing
            End If
        Catch ex As Exception
            Me.topo_category = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("expo_category"), String)) Then
                Me.expo_category = CType(SiteCodeDataRow.Item("expo_category"), String)
            Else
                Me.expo_category = Nothing
            End If
        Catch ex As Exception
            Me.expo_category = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("crest_height"), Double)) Then
                Me.crest_height = CType(SiteCodeDataRow.Item("crest_height"), Double)
            Else
                Me.crest_height = Nothing
            End If
        Catch ex As Exception
            Me.crest_height = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("slope_distance"), Double)) Then
                Me.slope_distance = CType(SiteCodeDataRow.Item("slope_distance"), Double)
            Else
                Me.slope_distance = Nothing
            End If
        Catch ex As Exception
            Me.slope_distance = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("distance_from_crest"), Double)) Then
                Me.distance_from_crest = CType(SiteCodeDataRow.Item("distance_from_crest"), Double)
            Else
                Me.distance_from_crest = Nothing
            End If
        Catch ex As Exception
            Me.distance_from_crest = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("downwind"), Boolean)) Then
                Me.downwind = CType(SiteCodeDataRow.Item("downwind"), Boolean)
            Else
                Me.downwind = Nothing
            End If
        Catch ex As Exception
            Me.downwind = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("topo_feature"), String)) Then
                Me.topo_feature = CType(SiteCodeDataRow.Item("topo_feature"), String)
            Else
                Me.topo_feature = Nothing
            End If
        Catch ex As Exception
            Me.topo_feature = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("crest_point_elev"), Double)) Then
                Me.crest_point_elev = CType(SiteCodeDataRow.Item("crest_point_elev"), Double)
            Else
                Me.crest_point_elev = Nothing
            End If
        Catch ex As Exception
            Me.crest_point_elev = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("base_point_elev"), Double)) Then
                Me.base_point_elev = CType(SiteCodeDataRow.Item("base_point_elev"), Double)
            Else
                Me.base_point_elev = Nothing
            End If
        Catch ex As Exception
            Me.base_point_elev = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("mid_height_elev"), Double)) Then
                Me.mid_height_elev = CType(SiteCodeDataRow.Item("mid_height_elev"), Double)
            Else
                Me.mid_height_elev = Nothing
            End If
        Catch ex As Exception
            Me.mid_height_elev = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("crest_to_mid_height_distance"), Double)) Then
                Me.crest_to_mid_height_distance = CType(SiteCodeDataRow.Item("crest_to_mid_height_distance"), Double)
            Else
                Me.crest_to_mid_height_distance = Nothing
            End If
        Catch ex As Exception
            Me.crest_to_mid_height_distance = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("tower_point_elev"), Double)) Then
                Me.tower_point_elev = CType(SiteCodeDataRow.Item("tower_point_elev"), Double)
            Else
                Me.tower_point_elev = Nothing
            End If
        Catch ex As Exception
            Me.tower_point_elev = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("base_kzt"), Double)) Then
                Me.base_kzt = CType(SiteCodeDataRow.Item("base_kzt"), Double)
            Else
                Me.base_kzt = Nothing
            End If
        Catch ex As Exception
            Me.base_kzt = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("site_name"), String)) Then
                Me.site_name = CType(SiteCodeDataRow.Item("site_name"), String)
            Else
                Me.site_name = Nothing
            End If
        Catch ex As Exception
            Me.site_name = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("structure_type"), String)) Then
                Me.structure_type = CType(SiteCodeDataRow.Item("structure_type"), String)
            Else
                Me.structure_type = Nothing
            End If
        Catch ex As Exception
            Me.structure_type = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("eng_app_id"), String)) Then
                Me.eng_app_id = CType(SiteCodeDataRow.Item("eng_app_id"), String)
            Else
                Me.eng_app_id = Nothing
            End If
        Catch ex As Exception
            Me.eng_app_id = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("eng_app_id_revision"), String)) Then
                Me.eng_app_id_revision = CType(SiteCodeDataRow.Item("eng_app_id_revision"), String)
            Else
                Me.eng_app_id_revision = Nothing
            End If
        Catch ex As Exception
            Me.eng_app_id_revision = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("lat_dec"), Double)) Then
                Me.lat_dec = CType(SiteCodeDataRow.Item("lat_dec"), Double)
            Else
                Me.lat_dec = Nothing
            End If
        Catch ex As Exception
            Me.lat_dec = Nothing
        End Try
        Try
            If Not IsDBNull(CType(SiteCodeDataRow.Item("long_dec"), Double)) Then
                Me.long_dec = CType(SiteCodeDataRow.Item("long_dec"), Double)
            Else
                Me.long_dec = Nothing
            End If
        Catch ex As Exception
            Me.long_dec = Nothing
        End Try
    End Sub
#End Region

End Class








