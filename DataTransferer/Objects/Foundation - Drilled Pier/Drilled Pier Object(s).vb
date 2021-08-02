Option Strict On
Imports System.ComponentModel

Partial Public Class DrilledPier

#Region "Define"
    Private prop_pier_id As Integer
    Private prop_local_pier_id As Integer
    Private prop_foundation_depth As Double
    Private prop_extension_above_grade As Double
    Private prop_groundwater_depth As Double
    Private prop_assume_min_steel As Boolean
    Private prop_check_shear_along_depth As Boolean
    Private prop_utilize_shear_friction_methodology As Boolean 'renamed
    Private prop_embedded_pole As Boolean
    Private prop_belled_pier As Boolean
    Private prop_soil_layer_quantity As Integer
    Private prop_concrete_compressive_strength As Double 'NEW
    Private prop_tie_yield_strength As Double 'NEW
    Private prop_longitudinal_rebar_yield_strength As Double 'NEW
    Private prop_rebar_effective_depths As Boolean 'NEW
    Private prop_rebar_cage_2_fy_override As Double 'NEW
    Private prop_rebar_cage_3_fy_override As Double 'NEW
    Private prop_shear_override_crit_depth As Boolean 'NEW
    Private prop_shear_crit_depth_override_comp As Double 'NEW
    Private prop_shear_crit_depth_override_uplift As Double 'NEW
    Private prop_bearing_type_toggle As Boolean 'MOVED
    Public Property soil_layers As New List(Of DrilledPierSoilLayer)
    Public Property sections As New List(Of DrilledPierSection)
    Public Property belled_details As DrilledPierBelledPier
    Public Property embed_details As DrilledPierEmbeddedPier

    <Category("Drilled Pier Details"), Description(""), DisplayName("Drilled Pier ID")>
    Public Property pier_id() As Integer
        Get
            Return Me.prop_pier_id
        End Get
        Set
            Me.prop_pier_id = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Drilled Pier ID")>
    Public Property local_pier_id() As Integer
        Get
            Return Me.prop_local_pier_id
        End Get
        Set
            Me.prop_local_pier_id = Value
        End Set
    End Property

    <Category("Drilled Pier Details"), Description(""), DisplayName("Foundation Depth")>
    Public Property foundation_depth() As Double
        Get
            Return Me.prop_foundation_depth
        End Get
        Set
            Me.prop_foundation_depth = Value
        End Set
    End Property

    <Category("Drilled Pier Details"), Description(""), DisplayName("Extension Above Grade")>
    Public Property extension_above_grade() As Double
        Get
            Return Me.prop_extension_above_grade
        End Get
        Set
            Me.prop_extension_above_grade = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Groundwater Depth")>
    Public Property groundwater_depth() As Double
        Get
            Return Me.prop_groundwater_depth
        End Get
        Set
            Me.prop_groundwater_depth = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Assume Minimum Steel")>
    Public Property assume_min_steel() As Boolean
        Get
            Return Me.prop_assume_min_steel
        End Get
        Set
            Me.prop_assume_min_steel = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Check Shear along Depth of Pier")>
    Public Property check_shear_along_depth() As Boolean
        Get
            Return Me.prop_check_shear_along_depth
        End Get
        Set
            Me.prop_check_shear_along_depth = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Utilize Shear-Friction Methodology")>
    Public Property utilize_shear_friction_methodology() As Boolean
        Get
            Return Me.prop_utilize_shear_friction_methodology
        End Get
        Set
            Me.prop_utilize_shear_friction_methodology = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Embedded Pole")>
    Public Property embedded_pole() As Boolean
        Get
            Return Me.prop_embedded_pole
        End Get
        Set
            Me.prop_embedded_pole = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Belled Pier")>
    Public Property belled_pier() As Boolean
        Get
            Return Me.prop_belled_pier
        End Get
        Set
            Me.prop_belled_pier = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Soil Layer Quantity")>
    Public Property soil_layer_quantity() As Integer
        Get
            Return Me.prop_soil_layer_quantity
        End Get
        Set
            Me.prop_soil_layer_quantity = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Concrete Compressive Strength")>
    Public Property concrete_compressive_strength() As Double
        Get
            Return Me.prop_concrete_compressive_strength
        End Get
        Set
            Me.prop_concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Tie Yield Strength")>
    Public Property tie_yield_strength() As Double
        Get
            Return Me.prop_tie_yield_strength
        End Get
        Set
            Me.prop_tie_yield_strength = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Longitudinal Rebar Yield Strength")>
    Public Property longitudinal_rebar_yield_strength() As Double
        Get
            Return Me.prop_longitudinal_rebar_yield_strength
        End Get
        Set
            Me.prop_longitudinal_rebar_yield_strength = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Rebar Effective Depths")>
    Public Property rebar_effective_depths() As Boolean
        Get
            Return Me.prop_rebar_effective_depths
        End Get
        Set
            Me.prop_rebar_effective_depths = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Rebar Cage 2 Fy Override")>
    Public Property rebar_cage_2_fy_override() As Double
        Get
            Return Me.prop_rebar_cage_2_fy_override
        End Get
        Set
            Me.prop_rebar_cage_2_fy_override = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Rebar Cage 3 Fy Override")>
    Public Property rebar_cage_3_fy_override() As Double
        Get
            Return Me.prop_rebar_cage_3_fy_override
        End Get
        Set
            Me.prop_rebar_cage_3_fy_override = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Override Shear Critical Depth")>
    Public Property shear_override_crit_depth() As Boolean
        Get
            Return Me.prop_shear_override_crit_depth
        End Get
        Set
            Me.prop_shear_override_crit_depth = False
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Shear Critical Depth Override - Compression")>
    Public Property shear_crit_depth_override_comp() As Double
        Get
            Return Me.prop_shear_crit_depth_override_comp
        End Get
        Set
            Me.prop_shear_crit_depth_override_comp = Value
        End Set
    End Property
    <Category("Drilled Pier Details"), Description(""), DisplayName("Shear Critical Depth Override - Uplift")>
    Public Property shear_crit_depth_override_uplift() As Double
        Get
            Return Me.prop_shear_crit_depth_override_uplift
        End Get
        Set
            Me.prop_shear_crit_depth_override_uplift = Value
        End Set
    End Property
    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Use Ultimate Bearing")>
    Public Property bearing_type_toggle() As Boolean
        Get
            Return Me.prop_bearing_type_toggle
        End Get
        Set
            Me.prop_bearing_type_toggle = Value
        End Set
    End Property
    '<Category("Drilled Pier Details"), Description(""), DisplayName("Soil Layers")>
    'Public Property soil_layers() As List(Of DrilledPierSoilLayer)
    '    Get
    '        Return Me.prop_soil_layers
    '    End Get
    '    Set
    '        Me.prop_soil_layers = Value
    '    End Set
    'End Property
    '<Category("Drilled Pier Details"), Description(""), DisplayName("Pier Sections")>
    'Public Property sections() As List(Of DrilledPierSection)
    '    Get
    '        Return Me.prop_sections
    '    End Get
    '    Set
    '        Me.prop_sections = Value
    '    End Set
    'End Property
    '<Category("Drilled Pier Details"), Description(""), DisplayName("Belled Pier Details")>
    'Public Property belled_details() As DrilledPierBelledPier
    '    Get
    '        Return Me.prop_belled_details
    '    End Get
    '    Set
    '        Me.prop_belled_details = Value
    '    End Set
    'End Property
    '<Category("Drilled Pier Details"), Description(""), DisplayName("Embedded Pier Details")>
    'Public Property embed_details() As DrilledPierEmbeddedPier
    '    Get
    '        Return Me.prop_embed_details
    '    End Get
    '    Set
    '        Me.prop_embed_details = Value
    '    End Set
    'End Property
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal DrilledPierDataRow As DataRow, refID As Integer)
        'General Drilled Pier Details
        Try
            Me.pier_id = CType(DrilledPierDataRow.Item("drilled_pier_id"), Integer)
        Catch
            Me.pier_id = 0
        End Try 'Drilled Pier ID
        Try
            Me.foundation_depth = CType(DrilledPierDataRow.Item("foundation_depth"), Double)
        Catch
            Me.foundation_depth = 0
        End Try 'Foundation Depth
        Try
            Me.extension_above_grade = CType(DrilledPierDataRow.Item("extension_above_grade"), Double)
        Catch
            Me.extension_above_grade = 0
        End Try 'Extension Above Grade
        Try
            Me.groundwater_depth = CType(DrilledPierDataRow.Item("groundwater_depth"), Double)
        Catch
            Me.groundwater_depth = 0
        End Try 'Groundwater Depth
        Try
            Me.assume_min_steel = CType(DrilledPierDataRow.Item("assume_min_steel"), Boolean)
        Catch
            Me.assume_min_steel = False
        End Try 'Assume Minimum Steel
        Try
            Me.check_shear_along_depth = CType(DrilledPierDataRow.Item("check_shear_along_depth"), Boolean)
        Catch
            Me.check_shear_along_depth = False
        End Try 'Check Shear along Depth of Pier
        Try
            Me.utilize_shear_friction_methodology = CType(DrilledPierDataRow.Item("utilize_shear_friction_methodology"), Boolean)
        Catch
            Me.utilize_shear_friction_methodology = False
        End Try 'Utilize Shear - Friction Methodology
        Try
            Me.embedded_pole = CType(DrilledPierDataRow.Item("embedded_pole"), Boolean)
        Catch
            Me.embedded_pole = False
        End Try 'Embedded Pole 
        Try
            Me.belled_pier = CType(DrilledPierDataRow.Item("belled_pier"), Boolean)
        Catch
            Me.belled_pier = False
        End Try 'Belled Pier
        Try
            Me.soil_layer_quantity = CType(DrilledPierDataRow.Item("soil_layer_quantity"), Integer)
        Catch
            Me.soil_layer_quantity = 0
        End Try 'Soil Layer Quantity
        Try
            Me.concrete_compressive_strength = CType(DrilledPierDataRow.Item("concrete_compressive_strength"), Double)
        Catch
            Me.concrete_compressive_strength = 0
        End Try 'Concrete Compressive Strength
        Try
            Me.tie_yield_strength = CType(DrilledPierDataRow.Item("tie_yield_strength"), Double)
        Catch
            Me.tie_yield_strength = 0
        End Try 'Tie Yield Strength
        Try
            Me.longitudinal_rebar_yield_strength = CType(DrilledPierDataRow.Item("longitudinal_rebar_yield_strength"), Double)
        Catch
            Me.longitudinal_rebar_yield_strength = 0
        End Try 'longitudinal_rebar_yield_strength
        Try
            Me.rebar_effective_depths = CType(DrilledPierDataRow.Item("rebar_effective_depths"), Boolean)
        Catch
            Me.rebar_effective_depths = True
        End Try 'rebar_effective_depths
        Try
            Me.rebar_cage_2_fy_override = CType(DrilledPierDataRow.Item("rebar_cage_2_fy_override"), Double)
        Catch
            Me.rebar_cage_2_fy_override = 0
        End Try 'rebar_cage_2_fy_override
        Try
            Me.rebar_cage_3_fy_override = CType(DrilledPierDataRow.Item("rebar_cage_3_fy_override"), Double)
        Catch
            Me.rebar_cage_3_fy_override = 0
        End Try 'rebar_cage_3_fy_override
        Try
            Me.bearing_type_toggle = CType(DrilledPierDataRow.Item("bearing_type_toggle"), Boolean)
        Catch
            Me.bearing_type_toggle = False
        End Try 'Use Ultimate Bearing

        For Each SoilLayerDataRow As DataRow In ds.Tables("Drilled Pier Soil SQL").Rows
            Dim soilRefID As Integer = CType(SoilLayerDataRow.Item("drilled_pier_id"), Integer)

            If soilRefID = refID Then
                Me.soil_layers.Add(New DrilledPierSoilLayer(SoilLayerDataRow))
            End If
        Next 'Add Soild Layers to to Drilled Pier Object

        For Each SectionDataRow As DataRow In ds.Tables("Drilled Pier Section SQL").Rows
            Dim secRefID As Integer = CType(SectionDataRow.Item("drilled_pier_id"), Integer)
            Dim secID As Integer = CType(SectionDataRow.Item("section_id"), Integer)

            If secRefID = refID Then
                Dim newSec As DrilledPierSection
                newSec = New DrilledPierSection(SectionDataRow)

                For Each RebarDataRow As DataRow In ds.Tables("Drilled Pier Rebar SQL").Rows
                    Dim rebSecID As Integer = CType(RebarDataRow.Item("section_id"), Integer)

                    If rebSecID = secID Then
                        newSec.rebar.Add(New DrilledPierRebar(RebarDataRow))
                    End If
                Next 'Add Drilled Pier Rebar to Section Object

                Me.sections.Add(newSec)
            End If
        Next 'Add Drilled Pier Sections to Drilled Pier Object

        If ds.Tables("Belled Details SQL").Rows.Count > 0 Then
            For Each BelledDataRow As DataRow In ds.Tables("Belled Details SQL").Rows
                Dim bellRefID As Integer = CType(BelledDataRow.Item("drilled_pier_id"), Integer)

                If bellRefID = refID Then
                    Me.belled_details = New DrilledPierBelledPier(BelledDataRow)

                    Exit For
                End If
            Next
        End If 'Add Belled Pier Details to Drilled Pier Object

        If ds.Tables("Embedded Details SQL").Rows.Count > 0 Then
            For Each EmbeddedDataRow As DataRow In ds.Tables("Embedded Details SQL").Rows
                Dim embedRefID As Integer = CType(EmbeddedDataRow.Item("drilled_pier_id"), Integer)
                Dim embedID As Integer = CType(EmbeddedDataRow.Item("embedded_id"), Integer)

                If embedRefID = refID Then
                    Me.embed_details = New DrilledPierEmbeddedPier(EmbeddedDataRow)

                    'For Each EmbedSectionDataRow As DataRow In ds.Tables("Embedded Section SQL").Rows
                    '    Dim eSec As New DrilledPierEmbedSection
                    '    Dim secEmbedID As Integer = CType(EmbedSectionDataRow.Item("embedded_id"), Integer)

                    '    If secEmbedID = embedID Then
                    '        Me.embed_details.sections.Add(New DrilledPierEmbedSection(EmbedSectionDataRow))
                    '    End If
                    'Next 'Add Embedded Pole Sections to embedded pole object

                    Exit For
                End If
            Next
        End If 'Add Embedded Pole Details to Drilled Pier Object
    End Sub 'Generate a drilled pier from EDS

    Public Sub New(ByVal DrilledPierDataRow As DataRow, ByVal refID As Integer, ByVal refcol As String)
        'General Drilled Pier Details
        Try
            Me.pier_id = CType(DrilledPierDataRow.Item("drilled_pier_id"), Integer)
        Catch
            Me.pier_id = 0
        End Try 'Drilled Pier ID
        Try
            Me.local_pier_id = CType(DrilledPierDataRow.Item("local_drilled_pier_id"), Integer)
        Catch
            Me.local_pier_id = 0
        End Try 'Local Drilled Pier ID
        Try
            Me.foundation_depth = CType(DrilledPierDataRow.Item("foundation_depth"), Double)
        Catch
            Me.foundation_depth = 0
        End Try 'Foundation Depth
        Try
            Me.extension_above_grade = CType(DrilledPierDataRow.Item("extension_above_grade"), Double)
        Catch
            Me.extension_above_grade = 0
        End Try 'Extension Above Grade
        Try
            Me.groundwater_depth = CType(DrilledPierDataRow.Item("groundwater_depth"), Double)
        Catch
            Me.groundwater_depth = 0
        End Try 'Groundwater Depth
        Try
            Me.assume_min_steel = CType(DrilledPierDataRow.Item("assume_min_steel"), Boolean)
        Catch
            Me.assume_min_steel = False
        End Try 'Assume Minimum Steel
        Try
            Me.check_shear_along_depth = CType(DrilledPierDataRow.Item("check_shear_along_depth"), Boolean)
        Catch
            Me.check_shear_along_depth = False
        End Try 'Check Shear along Depth of Pier
        Try
            Me.utilize_shear_friction_methodology = CType(DrilledPierDataRow.Item("utilize_shear_friction_methodology"), Boolean)
        Catch
            Me.utilize_shear_friction_methodology = False
        End Try 'Utilize Shear - Friction Methodology
        Try
            Me.embedded_pole = CType(DrilledPierDataRow.Item("embedded_pole"), Boolean)
        Catch
            Me.embedded_pole = False
        End Try 'Embedded Pole 
        Try
            Me.belled_pier = CType(DrilledPierDataRow.Item("belled_pier"), Boolean)
        Catch
            Me.belled_pier = False
        End Try 'Belled Pier
        Try
            Me.soil_layer_quantity = CType(DrilledPierDataRow.Item("soil_layer_quantity"), Integer)
        Catch
            Me.soil_layer_quantity = 0
        End Try 'Soil Layer Quantity

        For Each SoilLayerDataRow As DataRow In ds.Tables("Drilled Pier Soil EXCEL").Rows
            Dim soilRefID As Integer = CType(SoilLayerDataRow.Item(refcol), Integer)

            If soilRefID = refID Then
                Me.soil_layers.Add(New DrilledPierSoilLayer(SoilLayerDataRow))
            End If
        Next 'Add Soild Layers to to Drilled Pier Object

        For Each SectionDataRow As DataRow In ds.Tables("Drilled Pier Section EXCEL").Rows
            Dim secRefID As Integer = CType(SectionDataRow.Item(refcol), Integer)
            Dim secID As Integer = CType(SectionDataRow.Item("section_id"), Integer)

            If secRefID = refID Then
                Dim newSec As DrilledPierSection
                newSec = New DrilledPierSection(SectionDataRow)

                For Each RebarDataRow As DataRow In ds.Tables("Drilled Pier Rebar EXCEL").Rows
                    Dim rebSecID As Integer = CType(RebarDataRow.Item("section_id"), Integer)

                    If rebSecID = secID Then
                        newSec.rebar.Add(New DrilledPierRebar(RebarDataRow))
                    End If
                Next 'Add Drilled Pier Rebar to Section Object

                Me.sections.Add(newSec)
            End If
        Next 'Add Drilled Pier Sections to Drilled Pier Object

        If ds.Tables("Belled Details EXCEL").Rows.Count > 0 Then
            For Each BelledDataRow As DataRow In ds.Tables("Belled Details EXCEL").Rows
                Dim bellRefID As Integer = CType(BelledDataRow.Item(refcol), Integer)

                If bellRefID = refID Then
                    Me.belled_details = New DrilledPierBelledPier(BelledDataRow)

                    Exit For
                End If
            Next
        End If 'Add Belled Pier Details to Drilled Pier Object

        If ds.Tables("Embedded Details EXCEL").Rows.Count > 0 Then
            For Each EmbeddedDataRow As DataRow In ds.Tables("Embedded Details EXCEL").Rows
                Dim embedRefID As Integer = CType(EmbeddedDataRow.Item(refcol), Integer)
                Dim embedID As Integer = CType(EmbeddedDataRow.Item("embedded_id"), Integer)

                If embedRefID = refID Then
                    Me.embed_details = New DrilledPierEmbeddedPier(EmbeddedDataRow)

                    'For Each EmbedSectionDataRow As DataRow In ds.Tables("Embedded Section EXCEL").Rows
                    '    Dim eSec As New DrilledPierEmbedSection
                    '    Dim secEmbedID As Integer = CType(EmbedSectionDataRow.Item("embedded_id"), Integer)

                    '    If secEmbedID = embedID Then
                    '        Me.embed_details.sections.Add(New DrilledPierEmbedSection(EmbedSectionDataRow))
                    '    End If
                    'Next 'Add Embedded Pole Sections to embedded pole object

                    Exit For
                End If
            Next
        End If 'Add Embedded Pole Details to Drilled Pier Object
    End Sub 'Generate a drilled pier from Excel
#End Region

End Class


#Region "Drilled Pier Extras"
Partial Public Class DrilledPierSection
    Private prop_section_id As Integer
    Private prop_local_section_id As Integer
    Private prop_pier_diameter As Double
    Private prop_clear_cover As Double
    Private prop_clear_cover_rebar_cage_option As Boolean
    Private prop_tie_size As Integer
    Private prop_tie_spacing As Double
    'Private prop_top_elevation As Double
    Private prop_bottom_elevation As Double
    'Private prop_tie_yield_strength As Double
    'Private prop_concrete_compressive_strength As Double
    Private prop_assume_min_steel_rho_override As Double
    Public Property rebar As New List(Of DrilledPierRebar)

    <Category("Drilled Pier Sections"), Description(""), DisplayName("Section ID")>
    Public Property section_id() As Integer
        Get
            Return Me.prop_section_id
        End Get
        Set
            Me.prop_section_id = Value
        End Set
    End Property
    <Category("Drilled Pier Sections"), Description(""), DisplayName("Local Section ID")>
    Public Property local_section_id() As Integer
        Get
            Return Me.prop_local_section_id
        End Get
        Set
            Me.prop_local_section_id = Value
        End Set
    End Property
    <Category("Drilled Pier Sections"), Description(""), DisplayName("Pier Diameter")>
    Public Property pier_diameter() As Double
        Get
            Return Me.prop_pier_diameter
        End Get
        Set
            Me.prop_pier_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier Sections"), Description(""), DisplayName("Clear Cover")>
    Public Property clear_cover() As Double
        Get
            Return Me.prop_clear_cover
        End Get
        Set
            Me.prop_clear_cover = Value
        End Set
    End Property
    <Category("Drilled Pier Sections"), Description(""), DisplayName("")>
    Public Property clear_cover_rebar_cage_option() As Boolean
        Get
            Return Me.prop_clear_cover_rebar_cage_option
        End Get
        Set
            Me.prop_clear_cover_rebar_cage_option = Value
        End Set
    End Property
    <Category("Drilled Pier Sections"), Description(""), DisplayName("Tie Size")>
    Public Property tie_size() As Integer
        Get
            Return Me.prop_tie_size
        End Get
        Set
            Me.prop_tie_size = Value
        End Set
    End Property
    <Category("Drilled Pier Sections"), Description(""), DisplayName("Tie Spacing")>
    Public Property tie_spacing() As Double
        Get
            Return Me.prop_tie_spacing
        End Get
        Set
            Me.prop_tie_spacing = Value
        End Set
    End Property
    '<Category("Drilled Pier Sections"), Description(""), DisplayName("Top Elevation")> 'not utilized in tool. Calculated value in tool
    'Public Property top_elevation() As Double
    '    Get
    '        Return Me.prop_top_elevation
    '    End Get
    '    Set
    '        Me.prop_top_elevation = Value
    '    End Set
    'End Property
    <Category("Drilled Pier Sections"), Description(""), DisplayName("Bottom Elevation")>
    Public Property bottom_elevation() As Double
        Get
            Return Me.prop_bottom_elevation
        End Get
        Set
            Me.prop_bottom_elevation = Value
        End Set
    End Property
    '<Category("Drilled Pier Sections"), Description(""), DisplayName("Tie yield Strength")> 'moved to pier details
    'Public Property tie_yield_strength() As Double
    '    Get
    '        Return Me.prop_tie_yield_strength
    '    End Get
    '    Set
    '        Me.prop_tie_yield_strength = Value
    '    End Set
    'End Property
    '<Category("Drilled Pier Sections"), Description(""), DisplayName("Concrete Compressive Strength")>'moved to pier details
    'Public Property concrete_compressive_strength() As Double
    '    Get
    '        Return Me.prop_concrete_compressive_strength
    '    End Get
    '    Set
    '        Me.prop_concrete_compressive_strength = Value
    '    End Set
    'End Property
    <Category("Drilled Pier Sections"), Description(""), DisplayName("Minimum Steel Rho Override")>
    Public Property assume_min_steel_rho_override() As Double
        Get
            Return Me.prop_assume_min_steel_rho_override
        End Get
        Set
            Me.prop_assume_min_steel_rho_override = Value
        End Set
    End Property
    '<Category("Drilled Pier Sections"), Description(""), DisplayName("Rebar")>
    'Public Property rebar() As List(Of DrilledPierRebar)
    '    Get
    '        Return Me.prop_rebar
    '    End Get
    '    Set
    '        Me.prop_rebar = Value
    '    End Set
    'End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal SectionDataRow As DataRow)
        Try
            Me.section_id = CType(SectionDataRow.Item("section_id"), Integer)
        Catch
            Me.section_id = 0
        End Try 'Section ID
        Try
            Me.local_section_id = CType(SectionDataRow.Item("local_section_id"), Integer)
        Catch
            Me.local_section_id = 0
        End Try 'Local Section ID
        Try
            Me.pier_diameter = CType(SectionDataRow.Item("pier_diameter"), Double)
        Catch
            Me.pier_diameter = 0
        End Try 'Pier Diameter
        Try
            Me.clear_cover = CType(SectionDataRow.Item("clear_cover"), Double)
        Catch
            Me.clear_cover = 0
        End Try 'Clear Cover
        Try
            Me.clear_cover_rebar_cage_option = CType(SectionDataRow.Item("clear_cover_rebar_cage_option"), Boolean)
        Catch
            Me.clear_cover_rebar_cage_option = False
        End Try 'Rbar Cage Option 
        Try
            Me.tie_size = CType(SectionDataRow.Item("tie_size"), Integer)
        Catch
            Me.tie_size = 0
        End Try 'Tie Size
        Try
            Me.tie_spacing = CType(SectionDataRow.Item("tie_spacing"), Double)
        Catch
            Me.tie_spacing = 0
        End Try 'Tie Spacing
        'Try
        '    Me.top_elevation = CType(SectionDataRow.Item("top_elevation"), Double)
        'Catch
        '    Me.top_elevation = 0
        'End Try 'Top Elevation
        Try
            Me.bottom_elevation = CType(SectionDataRow.Item("bottom_elevation"), Double)
        Catch
            Me.bottom_elevation = 0
        End Try 'Bottom Elevation
        'Try
        '    Me.tie_yield_strength = CType(SectionDataRow.Item("tie_yield_strength"), Double)
        'Catch
        '    Me.tie_yield_strength = 0
        'End Try 'Tie yield Strength
        'Try
        '    Me.concrete_compressive_strength = CType(SectionDataRow.Item("concrete_compressive_strength"), Double)
        'Catch
        '    Me.concrete_compressive_strength = 0
        'End Try 'Concrete Compressive Strength
        Try
            Me.assume_min_steel_rho_override = CType(SectionDataRow.Item("assume_min_steel_rho_override"), Double)
        Catch
            Me.assume_min_steel_rho_override = 0
        End Try 'Minimum Steel Rho Override

    End Sub 'Add a Section to a Drilled Pier

End Class

Partial Public Class DrilledPierRebar
    Private prop_rebar_id As Integer
    Private prop_longitudinal_rebar_quantity As Integer
    Private prop_longitudinal_rebar_size As Integer
    Private prop_longitudinal_rebar_cage_diameter As Double
    Private prop_longitudinal_rebar_yield_strength As Double

    <Category("Drilled Pier Rebar"), Description(""), DisplayName("Rebar ID")>
    Public Property rebar_id() As Integer
        Get
            Return Me.prop_rebar_id
        End Get
        Set
            Me.prop_rebar_id = Value
        End Set
    End Property
    <Category("Drilled Pier Rebar"), Description(""), DisplayName("Rebar Quantity")>
    Public Property longitudinal_rebar_quantity() As Integer
        Get
            Return Me.prop_longitudinal_rebar_quantity
        End Get
        Set
            Me.prop_longitudinal_rebar_quantity = Value
        End Set
    End Property
    <Category("Drilled Pier Rebar"), Description(""), DisplayName("Rebar Size")>
    Public Property longitudinal_rebar_size() As Integer
        Get
            Return Me.prop_longitudinal_rebar_size
        End Get
        Set
            Me.prop_longitudinal_rebar_size = Value
        End Set
    End Property
    <Category("Drilled Pier Rebar"), Description(""), DisplayName("Cage Diameter")>
    Public Property longitudinal_rebar_cage_diameter() As Double
        Get
            Return Me.prop_longitudinal_rebar_cage_diameter
        End Get
        Set
            Me.prop_longitudinal_rebar_cage_diameter = Value
        End Set
    End Property
    <Category("Drilled Pier Rebar"), Description(""), DisplayName("Rebar Yield Strength")>
    Public Property longitudinal_rebar_yield_strength() As Double
        Get
            Return Me.prop_longitudinal_rebar_yield_strength
        End Get
        Set
            Me.prop_longitudinal_rebar_yield_strength = Value
        End Set
    End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal RebarDataRow As DataRow)
        Try
            Me.rebar_id = CType(RebarDataRow.Item("rebar_id"), Integer)
        Catch
            Me.rebar_id = 0
        End Try 'Rebar ID
        Try
            Me.longitudinal_rebar_quantity = CType(RebarDataRow.Item("longitudinal_rebar_quantity"), Integer)
        Catch
            Me.longitudinal_rebar_quantity = 0
        End Try 'Rebar Quantity
        Try
            Me.longitudinal_rebar_size = CType(RebarDataRow.Item("longitudinal_rebar_size"), Integer)
        Catch
            Me.longitudinal_rebar_size = 0
        End Try 'Rebar Size
        Try
            Me.longitudinal_rebar_cage_diameter = CType(RebarDataRow.Item("longitudinal_rebar_cage_diameter"), Double)
        Catch
            Me.longitudinal_rebar_cage_diameter = 0
        End Try 'Cage Diameter
        Try
            Me.longitudinal_rebar_yield_strength = CType(RebarDataRow.Item("longitudinal_rebar_yield_strength"), Double)
        Catch
            Me.longitudinal_rebar_yield_strength = 0
        End Try 'Rebar Yield Strength
    End Sub 'Add Rebar to a Section
End Class

Partial Public Class DrilledPierSoilLayer
    Private prop_soil_layer_id As Integer
    Private prop_bottom_depth As Double
    Private prop_effective_soil_density As Double
    Private prop_cohesion As Double
    Private prop_friction_angle As Double
    Private prop_skin_friction_override_comp As Double
    Private prop_skin_friction_override_uplift As Double
    'Private prop_bearing_type_toggle As Boolean
    Private prop_nominal_bearing_capacity As Double
    Private prop_spt_blow_count As Integer

    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Soil Layer ID")>
    Public Property soil_layer_id() As Integer
        Get
            Return Me.prop_soil_layer_id
        End Get
        Set
            Me.prop_soil_layer_id = Value
        End Set
    End Property
    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Bottom Depth")>
    Public Property bottom_depth() As Double
        Get
            Return Me.prop_bottom_depth
        End Get
        Set
            Me.prop_bottom_depth = Value
        End Set
    End Property
    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Effective Soil Density")>
    Public Property effective_soil_density() As Double
        Get
            Return Me.prop_effective_soil_density
        End Get
        Set
            Me.prop_effective_soil_density = Value
        End Set
    End Property
    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Cohesion")>
    Public Property cohesion() As Double
        Get
            Return Me.prop_cohesion
        End Get
        Set
            Me.prop_cohesion = Value
        End Set
    End Property
    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Angle of Friction")>
    Public Property friction_angle() As Double
        Get
            Return Me.prop_friction_angle
        End Get
        Set
            Me.prop_friction_angle = Value
        End Set
    End Property
    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Ultimate Skin Friction (Comp)")>
    Public Property skin_friction_override_comp() As Double
        Get
            Return Me.prop_skin_friction_override_comp
        End Get
        Set
            Me.prop_skin_friction_override_comp = Value
        End Set
    End Property
    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Ultimate Skin Friction (Tens)")>
    Public Property skin_friction_override_uplift() As Double
        Get
            Return Me.prop_skin_friction_override_uplift
        End Get
        Set
            Me.prop_skin_friction_override_uplift = Value
        End Set
    End Property
    '<Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Use Ultimate Bearing")>
    'Public Property bearing_type_toggle() As Boolean
    '    Get
    '        Return Me.prop_bearing_type_toggle
    '    End Get
    '    Set
    '        Me.prop_bearing_type_toggle = Value
    '    End Set
    'End Property
    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("Bearing Capacity")>
    Public Property nominal_bearing_capacity() As Double
        Get
            Return Me.prop_nominal_bearing_capacity
        End Get
        Set
            Me.prop_nominal_bearing_capacity = Value
        End Set
    End Property
    <Category("Drilled Pier Soil Layers"), Description(""), DisplayName("SPT Blow Count")>
    Public Property spt_blow_count() As Integer
        Get
            Return Me.prop_spt_blow_count
        End Get
        Set
            Me.prop_spt_blow_count = Value
        End Set
    End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal SoilLayerDataRow As DataRow)
        Try
            Me.soil_layer_id = CType(SoilLayerDataRow.Item("soil_layer_id"), Integer)
        Catch
            Me.soil_layer_id = 0
        End Try 'Soil Layer ID
        Try
            Me.bottom_depth = CType(SoilLayerDataRow.Item("bottom_depth"), Double)
        Catch
            Me.bottom_depth = 0
        End Try 'Bottom Depth
        Try
            Me.effective_soil_density = CType(SoilLayerDataRow.Item("effective_soil_density"), Double)
        Catch
            Me.effective_soil_density = 0
        End Try 'Effective Soil Density
        Try
            Me.cohesion = CType(SoilLayerDataRow.Item("cohesion"), Double)
        Catch
            Me.cohesion = 0
        End Try 'Cohesion
        Try
            Me.friction_angle = CType(SoilLayerDataRow.Item("friction_angle"), Double)
        Catch
            Me.friction_angle = 0
        End Try 'Angle of Friction
        Try
            Me.skin_friction_override_comp = CType(SoilLayerDataRow.Item("skin_friction_override_comp"), Double)
        Catch
            Me.skin_friction_override_comp = 0
        End Try 'Ultimate Skin Friction (Comp)
        Try
            Me.skin_friction_override_uplift = CType(SoilLayerDataRow.Item("skin_friction_override_uplift"), Double)
        Catch
            Me.skin_friction_override_uplift = 0
        End Try 'Ultimate Skin Friction (Tens)
        'Try
        '    Me.bearing_type_toggle = CType(SoilLayerDataRow.Item("bearing_type_toggle"), Boolean)
        'Catch
        '    Me.bearing_type_toggle = False
        'End Try 'Use Ultimate Bearing
        Try
            Me.nominal_bearing_capacity = CType(SoilLayerDataRow.Item("nominal_bearing_capacity"), Double)
        Catch
            Me.nominal_bearing_capacity = 0
        End Try 'Bearing Capacity
        Try
            Me.spt_blow_count = CType(SoilLayerDataRow.Item("spt_blow_count"), Integer)
        Catch
            Me.spt_blow_count = 0
        End Try 'SPT Blow Count
    End Sub 'Add a Soil Layer to a Drilled Pier

End Class

Partial Public Class DrilledPierBelledPier
    Private prop_belled_pier_id As Integer
    Private prop_belled_pier_option As Boolean
    Private prop_bottom_diameter_of_bell As Double
    Private prop_bell_input_type As String
    Private prop_bell_angle As Double
    Private prop_bell_height As Double
    Private prop_bell_toe_height As Double
    Private prop_neglect_top_soil_layer As Boolean
    Private prop_swelling_expansive_soil As Boolean
    Private prop_depth_of_expansive_soil As Double
    Private prop_expansive_soil_force As Double

    <Category("Belled Pier Details"), Description(""), DisplayName("Belled Pier ID")>
    Public Property belled_pier_id() As Integer
        Get
            Return Me.prop_belled_pier_id
        End Get
        Set
            Me.prop_belled_pier_id = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Belled Pier")>
    Public Property belled_pier_option() As Boolean
        Get
            Return Me.prop_belled_pier_option
        End Get
        Set
            Me.prop_belled_pier_option = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Bottom Diameter of Bell")>
    Public Property bottom_diameter_of_bell() As Double
        Get
            Return Me.prop_bottom_diameter_of_bell
        End Get
        Set
            Me.prop_bottom_diameter_of_bell = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Bell Input Type")>
    Public Property bell_input_type() As String
        Get
            Return Me.prop_bell_input_type
        End Get
        Set
            Me.prop_bell_input_type = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Bell Angle")>
    Public Property bell_angle() As Double
        Get
            Return Me.prop_bell_angle
        End Get
        Set
            Me.prop_bell_angle = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Bell Height")>
    Public Property bell_height() As Double
        Get
            Return Me.prop_bell_height
        End Get
        Set
            Me.prop_bell_height = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Bell Toe Height")>
    Public Property bell_toe_height() As Double
        Get
            Return Me.prop_bell_toe_height
        End Get
        Set
            Me.prop_bell_toe_height = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Neglect Top Soil Layer")>
    Public Property neglect_top_soil_layer() As Boolean
        Get
            Return Me.prop_neglect_top_soil_layer
        End Get
        Set
            Me.prop_neglect_top_soil_layer = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Swelling Expansive Soil")>
    Public Property swelling_expansive_soil() As Boolean
        Get
            Return Me.prop_swelling_expansive_soil
        End Get
        Set
            Me.prop_swelling_expansive_soil = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Depth of Expansive Soil")>
    Public Property depth_of_expansive_soil() As Double
        Get
            Return Me.prop_depth_of_expansive_soil
        End Get
        Set
            Me.prop_depth_of_expansive_soil = Value
        End Set
    End Property
    <Category("Belled Pier Details"), Description(""), DisplayName("Expansive Soil Force")>
    Public Property expansive_soil_force() As Double
        Get
            Return Me.prop_expansive_soil_force
        End Get
        Set
            Me.prop_expansive_soil_force = Value
        End Set
    End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal BelledDataRow As DataRow)
        Try
            Me.belled_pier_id = CType(BelledDataRow.Item("belled_pier_id"), Integer)
        Catch
            Me.belled_pier_id = 0
        End Try 'Belled Pier ID
        Try
            Me.belled_pier_option = CType(BelledDataRow.Item("belled_pier_option"), Boolean)
        Catch
            Me.belled_pier_option = False
        End Try 'Belled Pier
        Try
            Me.bottom_diameter_of_bell = CType(BelledDataRow.Item("bottom_diameter_of_bell"), Double)
        Catch
            Me.bottom_diameter_of_bell = 0
        End Try 'Bottom Diameter of Bell
        Try
            Me.bell_input_type = CType(BelledDataRow.Item("bell_input_type"), String)
        Catch
            Me.bell_input_type = ""
        End Try 'Bell Input Type
        Try
            Me.bell_angle = CType(BelledDataRow.Item("bell_angle"), Double)
        Catch
            Me.bell_angle = 0
        End Try 'Bell Angle
        Try
            Me.bell_height = CType(BelledDataRow.Item("bell_height"), Double)
        Catch
            Me.bell_height = 0
        End Try 'Bell Height
        Try
            Me.bell_toe_height = CType(BelledDataRow.Item("bell_toe_height"), Double)
        Catch
            Me.bell_toe_height = 0
        End Try 'Bell Toe Height
        Try
            Me.neglect_top_soil_layer = CType(BelledDataRow.Item("neglect_top_soil_layer"), Boolean)
        Catch
            Me.neglect_top_soil_layer = False
        End Try 'Neglect Top Soil Layer
        Try
            Me.swelling_expansive_soil = CType(BelledDataRow.Item("swelling_expansive_soil"), Boolean)
        Catch
            Me.swelling_expansive_soil = False
        End Try 'Swelling Expansive Soil
        Try
            Me.depth_of_expansive_soil = CType(BelledDataRow.Item("depth_of_expansive_soil"), Double)
        Catch
            Me.depth_of_expansive_soil = 0
        End Try 'Depth of Expansive Soil
        Try
            Me.expansive_soil_force = CType(BelledDataRow.Item("expansive_soil_force"), Double)
        Catch
            Me.expansive_soil_force = 0
        End Try 'Expansive Soil Force
    End Sub 'Add Belled Pier data to a Drilled Pier
End Class

Partial Public Class DrilledPierEmbeddedPier
    Private prop_embedded_id As Integer
    Private prop_embedded_pole_option As Boolean
    Private prop_encased_in_concrete As Boolean
    Private prop_pole_side_quantity As Integer
    Private prop_pole_yield_strength As Double
    Private prop_pole_thickness As Double
    Private prop_embedded_pole_input_type As String
    Private prop_pole_diameter_toc As Double
    Private prop_pole_top_diameter As Double
    Private prop_pole_bottom_diameter As Double
    Private prop_pole_section_length As Double
    Private prop_pole_taper_factor As Double
    Private prop_pole_bend_radius_override As Double
    Public Property sections As New List(Of DrilledPierEmbedSection)

    <Category("Embedded Pier Details"), Description(""), DisplayName("Embedded Pole ID")>
    Public Property embedded_id() As Integer
        Get
            Return Me.prop_embedded_id
        End Get
        Set
            Me.prop_embedded_id = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Embedded Pole")>
    Public Property embedded_pole_option() As Boolean
        Get
            Return Me.prop_embedded_pole_option
        End Get
        Set
            Me.prop_embedded_pole_option = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Encased in Concrete")>
    Public Property encased_in_concrete() As Boolean
        Get
            Return Me.prop_encased_in_concrete
        End Get
        Set
            Me.prop_encased_in_concrete = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Pole Sides")>
    Public Property pole_side_quantity() As Integer
        Get
            Return Me.prop_pole_side_quantity
        End Get
        Set
            Me.prop_pole_side_quantity = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Pole Yield Strength")>
    Public Property pole_yield_strength() As Double
        Get
            Return Me.prop_pole_yield_strength
        End Get
        Set
            Me.prop_pole_yield_strength = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Pole Thickness")>
    Public Property pole_thickness() As Double
        Get
            Return Me.prop_pole_thickness
        End Get
        Set
            Me.prop_pole_thickness = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Embedded Input Type")>
    Public Property embedded_pole_input_type() As String
        Get
            Return Me.prop_embedded_pole_input_type
        End Get
        Set
            Me.prop_embedded_pole_input_type = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Pole Diameter TOC")>
    Public Property pole_diameter_toc() As Double
        Get
            Return Me.prop_pole_diameter_toc
        End Get
        Set
            Me.prop_pole_diameter_toc = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Pole Top Diameter")>
    Public Property pole_top_diameter() As Double
        Get
            Return Me.prop_pole_top_diameter
        End Get
        Set
            Me.prop_pole_top_diameter = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Pole Bottom Diameter")>
    Public Property pole_bottom_diameter() As Double
        Get
            Return Me.prop_pole_bottom_diameter
        End Get
        Set
            Me.prop_pole_bottom_diameter = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Pole Section Length")>
    Public Property pole_section_length() As Double
        Get
            Return Me.prop_pole_section_length
        End Get
        Set
            Me.prop_pole_section_length = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Pole Taper Factor")>
    Public Property pole_taper_factor() As Double
        Get
            Return Me.prop_pole_taper_factor
        End Get
        Set
            Me.prop_pole_taper_factor = Value
        End Set
    End Property
    <Category("Embedded Pier Details"), Description(""), DisplayName("Pole Bend Radius Override")>
    Public Property pole_bend_radius_override() As Double
        Get
            Return Me.prop_pole_bend_radius_override
        End Get
        Set
            Me.prop_pole_bend_radius_override = Value
        End Set
    End Property
    '<Category("Embedded Pier Details"), Description(""), DisplayName("Embedded Pier Sections")>
    'Public Property sections() As List(Of DrilledPierEmbedSection)
    '    Get
    '        Return Me.prop_sections
    '    End Get
    '    Set
    '        Me.prop_sections = Value
    '    End Set
    'End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal EmbeddedDataRow As DataRow)
        Try
            Me.embedded_id = CType(EmbeddedDataRow.Item("embedded_id"), Integer)
        Catch
            Me.embedded_id = 0
        End Try 'Embedded Pole ID
        Try
            Me.embedded_pole_option = CType(EmbeddedDataRow.Item("embedded_pole_option"), Boolean)
        Catch
            Me.embedded_pole_option = False
        End Try 'Embedded Pole
        Try
            Me.encased_in_concrete = CType(EmbeddedDataRow.Item("encased_in_concrete"), Boolean)
        Catch
            Me.encased_in_concrete = False
        End Try 'Encased in Concrete
        Try
            Me.pole_side_quantity = CType(EmbeddedDataRow.Item("pole_side_quantity"), Integer)
        Catch
            Me.pole_side_quantity = 0
        End Try 'Pole Sides
        Try
            Me.pole_yield_strength = CType(EmbeddedDataRow.Item("pole_yield_strength"), Double)
        Catch
            Me.pole_yield_strength = 0
        End Try 'Pole Yield Strength
        Try
            Me.pole_thickness = CType(EmbeddedDataRow.Item("pole_thickness"), Double)
        Catch
            Me.pole_thickness = 0
        End Try 'Pole Thickness
        Try
            Me.embedded_pole_input_type = CType(EmbeddedDataRow.Item("embedded_pole_input_type"), String)
        Catch
            Me.embedded_pole_input_type = ""
        End Try 'Embedded Input Type
        Try
            Me.pole_diameter_toc = CType(EmbeddedDataRow.Item("pole_diameter_toc"), Double)
        Catch
            Me.pole_diameter_toc = 0
        End Try 'Pole Diameter TOC
        Try
            Me.pole_top_diameter = CType(EmbeddedDataRow.Item("pole_top_diameter"), Double)
        Catch
            Me.pole_top_diameter = 0
        End Try 'Pole Top Diameter
        Try
            Me.pole_bottom_diameter = CType(EmbeddedDataRow.Item("pole_bottom_diameter"), Double)
        Catch
            Me.pole_bottom_diameter = 0
        End Try 'Pole Bottom Diameter
        Try
            Me.pole_section_length = CType(EmbeddedDataRow.Item("pole_section_length"), Double)
        Catch
            Me.pole_section_length = 0
        End Try 'Pole Section Length
        Try
            Me.pole_taper_factor = CType(EmbeddedDataRow.Item("pole_taper_factor"), Double)
        Catch
            Me.pole_taper_factor = 0
        End Try 'Pole Taper Factor
        Try
            Me.pole_bend_radius_override = CType(EmbeddedDataRow.Item("pole_bend_radius_override"), Double)
        Catch
            Me.pole_bend_radius_override = 0
        End Try 'Pole Bend Radius Override

    End Sub 'Add Embedded Pole data to a Drilled Pier

End Class

Partial Public Class DrilledPierEmbedSection
    Private prop_section_id As Integer
    Private prop_pier_diameter As Double

    <Category("Embedded Pier Sections"), Description(""), DisplayName("Embedded Section ID")>
    Public Property section_id() As Integer
        Get
            Return Me.prop_section_id
        End Get
        Set
            Me.prop_section_id = Value
        End Set
    End Property
    <Category("Embedded Pier Sections"), Description(""), DisplayName("Pier Diameter")>
    Public Property pier_diameter() As Double
        Get
            Return Me.prop_pier_diameter
        End Get
        Set
            Me.prop_pier_diameter = Value
        End Set
    End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal EmbedSectionDataRow As DataRow)
        Try
            Me.section_id = CType(EmbedSectionDataRow.Item("embedded_section_ID"), Integer)
        Catch
            Me.section_id = 0
        End Try 'Section ID
        Try
            Me.pier_diameter = CType(EmbedSectionDataRow.Item("pier_diameter"), Double)
        Catch
            Me.pier_diameter = 0
        End Try 'Pier Diameter
    End Sub 'Added Embedded Pole Sections to an Embedded Pole object

End Class
#End Region

