'Public Class CCIplate
'    Public Property connections As New List(Of Connection)
'    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)
'        Me.property1 = CType(dr.Item("property1"), String)
'        Dim plConnection As Connection
'        Dim plPlate As Plate
'        Dim plBolt As Bolt
'        Dim plMaterial As Material
'        'Loop through all connections
'        For Each dr As DataRow In ds.Tables(plConnection.EDSTableName).Rows
'            'create a new connection based on the datarow from above
'            plConnection = New Connection(dr, Me)
'            'Check if the parent id, in the case cciplate id is equal to the original object id (Me)                    
'            If plConnection.cciplate_id = Me.ID Then
'                'If it is equal then add the newly created connection to the list of connections 
'                connections.Add(plConnection)
'                'Loop through all plates pulled from EDS and check if they are associated with the newly created connection
'                For Each drr As DataRow In ds.Tables(plPlate.EDSTableName).Rows
'                    'Create a new plate from the plate datarow from EDS
'                    plPlate = New Plate(drr, plConnection)
'                    If plPlate.connection_id = plConnection.ID Then
'                        For Each drrr As DataRow In ds.Tables(plMaterial.EDSTableName).Rows
'                            plMaterial = New Material(drrr, plPlate)
'                            If plMaterial.name = plPlate.material_name Then
'                                plPlate.material = plMaterial
'                            End If
'                        Next
'                        For Each drrr As DataRow In ds.Tables(plBolt.EDSTableName).Rows)
'                                    plBolt = New Bolt(drrr, plPlate)
'                            If plBolt.plate_id = plPlate.ID Then
'                                plPlate.bolts.Add(plBolt)
'                            End If
'                        Next
'                    End If
'                Next
'            End If
'        Next
'    End Sub
'End Class
'Public Class Connection
'    Public Property plates As New List(Of Plate)
'    Public Property results As New List(Of Result)
'    Public Sub New(ByVal dr As DataRow, Option ByVal Parent As EDSObjcect)
'        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
'        If Parent IsNot Nothing Then Me.Absorb(Parent)
'        Me.property1 = CType(dr.Item("property1"), String)
'        Me.property2 = CType(dr.Item("property2"), Integer)
'        Me.property3 = CType(dr.Item("property3"), String)
'    End Sub
'End Class
'Public Class Plate
'    Public Property bolts As New List(Of Bolt)
'    Public Property results As New List(Of Result)
'    Public Property material As Material
'    Public Property material_name As String
'    Public Property connection_id As Integer
'End Class
'Public Class Bolt
'    Public Property results As New List(Of Result)
'    Public Property plate_id As Integer
'End Class
'Public Class Material
'End Class
