Imports System.Runtime.Serialization

'All SiteInfo information comes from Oracle
<DataContractAttribute()>
Public Class SiteInfo

#Region "Constructors"
    Public Sub New()

    End Sub
    Public Sub New(wo As String)

        'Create object from query
        Using strDS As New DataSet
            Dim query = "
                        SELECT
                            wo.WORK_ORDER_SEQ_NUM wo            
                            ,wo.BUS_UNIT bu                           
                            ,wo.STRUCTURE_ID sid
    
                            ,wo.ENG_APP_ID   app_num
                            ,wo.crrnt_prjct_rvsn_num app_rev                                                           
                            ,si.SITE_NAME site_name             
                            ,ea.JDE_JOB_NUM jde_num         
                            ,str.MFG tower_man
                            ,str.structure_type
                            ,addr.address_1 site_add
                            ,addr.city site_city
                            ,addr.county site_county
                            ,addr.state site_state
    
                            ,ROUND(str.LAT_DEC,8) lat_decimal
                            ,ROUND(str.LONG_DEC,8) long_decimal

                            ,org.org_name

                            ,ea.cust_site_alias_name customer_site_name 
                            ,ea.cust_site_num customer_site_num 

                            ,ea.cstmr_pymnt_ref_text fa_num
    
                        FROM
                            isit_aim.work_orders wo
                            ,isit_aim.structure str
                            ,isit_aim.site si
                            ,isit_isite.eng_application ea
                            ,isit_aim.address addr
                            ,isit_aim.org org
    
                        WHERE
                            wo.bus_unit = str.bus_unit (+)
                            AND wo.structure_id = str.structure_id (+)
                            AND wo.bus_unit = si.bus_unit (+)
                            AND si.address_id=addr.address_id(+)
                            AND wo.eng_app_id = ea.eng_app_id (+)
                            AND addr.ctry_id='US'
                            AND ea.org_seq_num = org.org_seq_num (+)
                            AND wo.work_order_seq_num='" & wo & "'"


            Dim x As Boolean = OracleLoader(query, "Site Info", strDS, 3000, "ords")

            If strDS.Tables.Contains("Site Info") Then
                If strDS.Tables("Site Info").Rows.Count > 0 Then

                    Dim SiteCodeDataRow = strDS.Tables("Site Info").Rows(0)

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("wo"), String)) Then
                            Me.wo = CType(SiteCodeDataRow.Item("wo"), String)
                        Else
                            Me.wo = Nothing
                        End If
                    Catch ex As Exception
                        Me.wo = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("bu"), String)) Then
                            Me.bu_num = CType(SiteCodeDataRow.Item("bu"), String)
                        Else
                            Me.bu_num = Nothing
                        End If
                    Catch ex As Exception
                        Me.bu_num = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("sid"), String)) Then
                            Me.structure_id = CType(SiteCodeDataRow.Item("sid"), String)
                        Else
                            Me.structure_id = Nothing
                        End If
                    Catch ex As Exception
                        Me.structure_id = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("app_num"), String)) Then
                            Me.app_num = CType(SiteCodeDataRow.Item("app_num"), String)
                        Else
                            Me.app_num = Nothing
                        End If
                    Catch ex As Exception
                        Me.app_num = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("app_rev"), String)) Then
                            Me.app_rev = CType(SiteCodeDataRow.Item("app_rev"), String)
                        Else
                            Me.app_rev = Nothing
                        End If
                    Catch ex As Exception
                        Me.app_rev = Nothing
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
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("jde_num"), String)) Then
                            Me.jde_num = CType(SiteCodeDataRow.Item("jde_num"), String)
                        Else
                            Me.jde_num = Nothing
                        End If
                    Catch ex As Exception
                        Me.jde_num = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("tower_man"), String)) Then
                            Me.tower_man = CType(SiteCodeDataRow.Item("tower_man"), String)
                        Else
                            Me.tower_man = Nothing
                        End If
                    Catch ex As Exception
                        Me.tower_man = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("structure_type"), String)) Then
                            Me.tower_type = CType(SiteCodeDataRow.Item("structure_type"), String)
                        Else
                            Me.tower_type = Nothing
                        End If
                    Catch ex As Exception
                        Me.tower_type = Nothing
                    End Try


                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("site_add"), String)) Then
                            Me.site_add = CType(SiteCodeDataRow.Item("site_add"), String)
                        Else
                            Me.site_add = Nothing
                        End If
                    Catch ex As Exception
                        Me.site_add = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("site_city"), String)) Then
                            Me.site_city = CType(SiteCodeDataRow.Item("site_city"), String)
                        Else
                            Me.site_city = Nothing
                        End If
                    Catch ex As Exception
                        Me.site_city = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("site_county"), String)) Then
                            Me.site_county = CType(SiteCodeDataRow.Item("site_county"), String)
                        Else
                            Me.site_county = Nothing
                        End If
                    Catch ex As Exception
                        Me.site_county = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("site_state"), String)) Then
                            Me.site_state = CType(SiteCodeDataRow.Item("site_state"), String)
                        Else
                            Me.site_state = Nothing
                        End If
                    Catch ex As Exception
                        Me.site_state = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("lat_decimal"), String)) Then
                            Me.lat_decimal = CType(SiteCodeDataRow.Item("lat_decimal"), String)
                        Else
                            Me.lat_decimal = Nothing
                        End If
                    Catch ex As Exception
                        Me.lat_decimal = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("long_decimal"), String)) Then
                            Me.long_decimal = CType(SiteCodeDataRow.Item("long_decimal"), String)
                        Else
                            Me.long_decimal = Nothing
                        End If
                    Catch ex As Exception
                        Me.long_decimal = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("org_name"), String)) Then
                            Me.customer = CType(SiteCodeDataRow.Item("org_name"), String)
                        Else
                            Me.customer = Nothing
                        End If
                    Catch ex As Exception
                        Me.customer = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("customer_site_name"), String)) Then
                            Me.cust_site_name = CType(SiteCodeDataRow.Item("customer_site_name"), String)
                        Else
                            Me.cust_site_name = Nothing
                        End If
                    Catch ex As Exception
                        Me.cust_site_name = Nothing
                    End Try

                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("customer_site_num"), String)) Then
                            Me.cust_site_num = CType(SiteCodeDataRow.Item("customer_site_num"), String)
                        Else
                            Me.cust_site_num = Nothing
                        End If
                    Catch ex As Exception
                        Me.cust_site_num = Nothing
                    End Try
                    Try
                        If Not IsDBNull(CType(SiteCodeDataRow.Item("fa_num"), String)) Then
                            Me.fa_num = CType(SiteCodeDataRow.Item("fa_num"), String)
                        Else
                            Me.fa_num = Nothing
                        End If
                    Catch ex As Exception
                        Me.fa_num = Nothing
                    End Try

                End If
            End If

            'Dim DocumentDataRow = strDS.Tables("Documents").Rows(0)

        End Using
    End Sub

#End Region

#Region "Properties"
     <DataMember()> Public Property wo As String
     <DataMember()> Public Property bu_num As Integer
     <DataMember()> Public Property structure_id As String

     <DataMember()> Public Property app_num As String
     <DataMember()> Public Property app_rev As String
     <DataMember()> Public Property site_name As String
     <DataMember()> Public Property jde_num As String
     <DataMember()> Public Property tower_man As String
     <DataMember()> Public Property tower_type As String

     <DataMember()> Public Property site_add As String
     <DataMember()> Public Property site_city As String
     <DataMember()> Public Property site_county As String
     <DataMember()> Public Property site_state As String

     <DataMember()> Public Property customer As String
     <DataMember()> Public Property cust_site_num As String
     <DataMember()> Public Property cust_site_name As String


     <DataMember()> Public Property fa_num As String
#End Region

#Region "Lat Long calculations"
     <DataMember()> Public Property lat_decimal As Double
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


     <DataMember()> Public Property long_decimal As Double
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

#End Region

End Class
