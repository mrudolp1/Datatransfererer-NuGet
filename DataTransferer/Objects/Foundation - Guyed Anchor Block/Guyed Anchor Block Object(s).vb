Option Strict Off
Option Compare Binary

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
'Imports Microsoft.Office.Interop

Partial Public Class AnchorBlockFoundation
    Inherits EDSExcelObject

    Public Property GuyedAnchorBlocks As New List(Of AnchorBlock)

    'Origin row in the driled pier database. Basically just where the profile numbers are in the database worksheet.
    'This is actually 58 but due to the 0,0 origin in excel, it is 1 less
    Private pierProfileRow As Integer = 57

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Foundation"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_tool"
        End Get
    End Property

    Public Overrides ReadOnly Property templatePath As String
        Get
            Return IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Guyed Anchor Block Foundation.xlsm")
        End Get
    End Property

    Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
        Get
            Dim ab As New AnchorBlock
            Dim abProf As New AnchorBlockProfile
            Dim abSProf As New AnchorBlockSoilProfile
            Dim abSlay As New AnchorBlockSoilLayer
            Dim abRes As New AnchorBlockResult
            Dim abTool As New AnchorBlockFoundation

            Return New List(Of EXCELDTParameter) From {
                                                        New EXCELDTParameter(ab.EDSObjectName, "A2:H52", "Profiles (ENTER)"),  'It is slightly confusing but to keep naming issues consistent in the tool a drilled pier = profile and a drilled pier profile = drilled pier details
                                                        New EXCELDTParameter(abProf.EDSObjectName, "A2:X52", "Details (ENTER)"),
                                                        New EXCELDTParameter(abSProf.EDSObjectName, "A2:E52", "Soil Profiles (ENTER)"),
                                                        New EXCELDTParameter(abSlay.EDSObjectName, "A2:N1502", "Soil Layers (ENTER)"),
                                                        New EXCELDTParameter(abRes.EDSObjectName, "BD8:BV58", "Foundation Input"),
                                                        New EXCELDTParameter(abTool.EDSObjectName, "A1:E2", "Tool (RETURN_ENTER)")
                                                                                        }
            '***Add additional table references here****
        End Get
    End Property

    Public Overrides Function SQLInsertValues() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLInsertFields() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Throw New NotImplementedException()
    End Function

    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        Throw New NotImplementedException()
    End Sub
End Class

Partial Public Class AnchorBlock
    Inherits EDSObjectWithQueries

    Public Property PierProfile As AnchorBlockProfile
    Public Property SoilProfile As AnchorBlockSoilProfile

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Profile"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete
    End Function

#End Region

#Region "Define"
    Private _ID As Integer?
    Private _anchor_block_profile_id As Integer?
    Private _soil_profile_id As Integer?
    Private _reaction_position As Integer?
    Private _reaction_location As String
    Private _local_anchor_profile As Integer?
    Private _local_soil_profile As Integer?
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Anchor Block Profile Id")>
    Public Property anchor_block_profile_id() As Integer?
        Get
            Return Me._anchor_block_profile_id
        End Get
        Set
            Me._anchor_block_profile_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Soil Profile Id")>
    Public Property soil_profile_id() As Integer?
        Get
            Return Me._soil_profile_id
        End Get
        Set
            Me._soil_profile_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Reaction Position")>
    Public Property reaction_position() As Integer?
        Get
            Return Me._reaction_position
        End Get
        Set
            Me._reaction_position = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Reaction Location")>
    Public Property reaction_location() As String
        Get
            Return Me._reaction_location
        End Get
        Set
            Me._reaction_location = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Local Anchor Profile")>
    Public Property local_anchor_profile() As Integer?
        Get
            Return Me._local_anchor_profile
        End Get
        Set
            Me._local_anchor_profile = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profile"), Description(""), DisplayName("Local Soil Profile")>
    Public Property local_soil_profile() As Integer?
        Get
            Return Me._local_soil_profile
        End Get
        Set
            Me._local_soil_profile = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow)
        ConstructMe(dr)
    End Sub

    Public Sub ConstructMe(ByVal dr As DataRow)
    End Sub

    Public Sub New(ByVal dr As DataRow, ByVal strDS As DataSet, ByVal isExcel As Boolean, Optional ByRef Parent As EDSObject = Nothing)
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String

        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String

        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String

        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        Return Equals

    End Function

End Class

Partial Public Class AnchorBlockProfile
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_profile"
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String
        Return SQLInsert
    End Function

    Public Overrides Function SQLUpdate() As String
        Return SQLUpdate
    End Function

    Public Overrides Function SQLDelete() As String
        SQLDelete = CCI_Engineering_Templates.My.Resources.General__DELETE
        SQLDelete = SQLDelete.Replace("[TABLE]", Me.EDSTableName)
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID)

        Return SQLDelete
    End Function

#End Region

#Region "Define"
    Private _ID As Integer?
    Private _anchor_location As String
    Private _guy_anchor_radius As Double?
    Private _anchor_depth As Double?
    Private _anchor_width As Double?
    Private _anchor_thickness As Double?
    Private _anchor_length As Double?
    Private _anchor_toe_width As Double?
    Private _anchor_top_rebar_size As Integer?
    Private _anchor_top_rebar_quantity As Integer?
    Private _anchor_bottom_rebar_size As Integer?
    Private _anchor_bottom_rebar_quantity As Integer?
    Private _anchor_stirrup_size As Integer?
    Private _anchor_shaft_diameter As Double?
    Private _anchor_shaft_quantity As Integer?
    Private _anchor_shaft_area_override As Double?
    Private _anchor_shaft_shear_leg_factor As Double?
    Private _rebar_grade As Double?
    Private _concrete_compressive_strength As Double?
    Private _clear_cover As Double?
    Private _anchor_shaft_yield_strength As Double?
    Private _anchor_shaft_ultimate_strength As Double?
    Private _tool_version As String
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Location")>
    Public Property anchor_location() As String
        Get
            Return Me._anchor_location
        End Get
        Set
            Me._anchor_location = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guy Anchor Radius")>
    Public Property guy_anchor_radius() As Double?
        Get
            Return Me._guy_anchor_radius
        End Get
        Set
            Me._guy_anchor_radius = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Depth")>
    Public Property anchor_depth() As Double?
        Get
            Return Me._anchor_depth
        End Get
        Set
            Me._anchor_depth = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Width")>
    Public Property anchor_width() As Double?
        Get
            Return Me._anchor_width
        End Get
        Set
            Me._anchor_width = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Thickness")>
    Public Property anchor_thickness() As Double?
        Get
            Return Me._anchor_thickness
        End Get
        Set
            Me._anchor_thickness = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description("column - 6+local drilled pier id"), DisplayName("Anchor Length")>
    Public Property anchor_length() As Double?
        Get
            Return Me._anchor_length
        End Get
        Set
            Me._anchor_length = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Toe Width")>
    Public Property anchor_toe_width() As Double?
        Get
            Return Me._anchor_toe_width
        End Get
        Set
            Me._anchor_toe_width = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Top Rebar Size")>
    Public Property anchor_top_rebar_size() As Integer?
        Get
            Return Me._anchor_top_rebar_size
        End Get
        Set
            Me._anchor_top_rebar_size = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Top Rebar Quantity")>
    Public Property anchor_top_rebar_quantity() As Integer?
        Get
            Return Me._anchor_top_rebar_quantity
        End Get
        Set
            Me._anchor_top_rebar_quantity = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Bottom Rebar Size")>
    Public Property anchor_bottom_rebar_size() As Integer?
        Get
            Return Me._anchor_bottom_rebar_size
        End Get
        Set
            Me._anchor_bottom_rebar_size = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Bottom Rebar Quantity")>
    Public Property anchor_bottom_rebar_quantity() As Integer?
        Get
            Return Me._anchor_bottom_rebar_quantity
        End Get
        Set
            Me._anchor_bottom_rebar_quantity = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Stirrup Size")>
    Public Property anchor_stirrup_size() As Integer?
        Get
            Return Me._anchor_stirrup_size
        End Get
        Set
            Me._anchor_stirrup_size = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Diameter")>
    Public Property anchor_shaft_diameter() As Double?
        Get
            Return Me._anchor_shaft_diameter
        End Get
        Set
            Me._anchor_shaft_diameter = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Quantity")>
    Public Property anchor_shaft_quantity() As Integer?
        Get
            Return Me._anchor_shaft_quantity
        End Get
        Set
            Me._anchor_shaft_quantity = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Area Override")>
    Public Property anchor_shaft_area_override() As Double?
        Get
            Return Me._anchor_shaft_area_override
        End Get
        Set
            Me._anchor_shaft_area_override = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Shear Leg Factor")>
    Public Property anchor_shaft_shear_leg_factor() As Double?
        Get
            Return Me._anchor_shaft_shear_leg_factor
        End Get
        Set
            Me._anchor_shaft_shear_leg_factor = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Rebar Grade")>
    Public Property rebar_grade() As Double?
        Get
            Return Me._rebar_grade
        End Get
        Set
            Me._rebar_grade = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Concrete Compressive Strength")>
    Public Property concrete_compressive_strength() As Double?
        Get
            Return Me._concrete_compressive_strength
        End Get
        Set
            Me._concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Clear Cover")>
    Public Property clear_cover() As Double?
        Get
            Return Me._clear_cover
        End Get
        Set
            Me._clear_cover = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Yield Strength")>
    Public Property anchor_shaft_yield_strength() As Double?
        Get
            Return Me._anchor_shaft_yield_strength
        End Get
        Set
            Me._anchor_shaft_yield_strength = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Ultimate Strength")>
    Public Property anchor_shaft_ultimate_strength() As Double?
        Get
            Return Me._anchor_shaft_ultimate_strength
        End Get
        Set
            Me._anchor_shaft_ultimate_strength = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Tool Version")>
    Public Property tool_version() As String
        Get
            Return Me._tool_version
        End Get
        Set
            Me._tool_version = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByRef Parent As EDSObject = Nothing)
    End Sub
#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String

        Return SQLInsertValues

    End Function

    Public Overrides Function SQLInsertFields() As String

        Return SQLInsertFields

    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String

        Return SQLUpdateFieldsandValues

    End Function
#End Region

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        Return Equals

    End Function
End Class

Partial Public Class AnchorBlockSoilProfile
    Inherits SoilProfile

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Foundation"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_tool"
        End Get
    End Property

End Class

Partial Public Class AnchorBlockSoilLayer
    Inherits SoilLayer

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Foundation"
        End Get
    End Property

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_tool"
        End Get
    End Property

End Class

Partial Public Class AnchorBlockResult
    Inherits EDSResult

    Public ReadOnly Property EDSObjectName As String
        Get
            Return "Anchor Block Foundation"
        End Get
    End Property

    Public ReadOnly Property EDSTableName As String
        Get
            Return "fnd.anchor_block_tool"
        End Get
    End Property

End Class
