Imports MoreLinq

Partial Public Class tnxTowerOutput
    Public Sub ConverttoEDSResults(tnx As tnxGeometry)
        If Me.MemberCompression IsNot Nothing Then
            For Each section In Me.MemberCompression
                section.ConverttoEDSResults(tnx)
            Next
        End If

        If Me.MemberTension IsNot Nothing Then
            For Each section In Me.MemberTension
                section.ConverttoEDSResults(tnx)
            Next
        End If

        If Me.Guys IsNot Nothing Then
            For Each section In Me.Guys
                section.ConverttoEDSResults(tnx)
            Next
        End If

        If Me.BoltDesignData IsNot Nothing Then
            For Each section In Me.BoltDesignData
                section.ConverttoEDSResults(tnx, Me.MemberTension)
            Next
        End If

    End Sub
End Class

#Region "Member Compression"

Partial Public Class tnxTowerOutputMemberCompressionTowerSection
    Public Sub ConverttoEDSResults(tnx As tnxGeometry)
        Dim tnxSection As tnxGeometryRec = tnx.tnxSectionSelector(CInt(Me.Number))
        For Each compResultComponent In Me.ComponentType
            compResultComponent.AddMaxComponentResultstoSection(tnxSection)
        Next
    End Sub
End Class

Partial Public Class tnxTowerOutputMemberCompressionComponentType
    Public Function MaxMember() As tnxTowerOutputMemberCompressionMember
        Return Me.Member.MaxBy(Function(x) x.MaxRatio).FirstOrDefault
    End Function

    Public Sub AddMaxComponentResultstoSection(tnxSection As tnxGeometryRec)
        Dim controllingMember As tnxTowerOutputMemberCompressionMember = Me.MaxMember()
        'Compression
        tnxSection.TNXResults.Add(New tnxResult(String.Format("comp_{0}_p", Me.Name.ToLower),
                                             controllingMember.Compression.PDCRatio,
                                             controllingMember.Compression.phiPn,
                                             controllingMember.Compression.Pu,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Bending X
        tnxSection.TNXResults.Add(New tnxResult(String.Format("comp_{0}_mx", Me.Name.ToLower),
                                             controllingMember.Bending.MxDCRatio,
                                             controllingMember.Bending.phiMnx,
                                             controllingMember.Bending.Mux,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Bending Y
        tnxSection.TNXResults.Add(New tnxResult(String.Format("comp_{0}_my", Me.Name.ToLower),
                                             controllingMember.Bending.MyDCRatio,
                                             controllingMember.Bending.phiMny,
                                             controllingMember.Bending.Muy,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Shear
        tnxSection.TNXResults.Add(New tnxResult(String.Format("comp_{0}_v", Me.Name.ToLower),
                                             controllingMember.Shear.VDCRatio,
                                             controllingMember.Shear.phiVn,
                                             controllingMember.Shear.Vu,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Torsion
        tnxSection.TNXResults.Add(New tnxResult(String.Format("comp_{0}_t", Me.Name.ToLower),
                                             controllingMember.Shear.TDCRatio,
                                             controllingMember.Shear.phiTn,
                                             controllingMember.Shear.Tu,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Interaction
        tnxSection.TNXResults.Add(New tnxResult(String.Format("comp_{0}_int", Me.Name.ToLower),
                                             controllingMember.Interaction.CombinedDCRatio,
                                             Nothing,
                                             Nothing,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
    End Sub
End Class

Partial Public Class tnxTowerOutputMemberCompressionMember
    Public Function MaxRatio() As Double
        Return {Compression.PDCRatio, Bending.MxDCRatio, Bending.MyDCRatio, Shear.VDCRatio, Shear.TDCRatio, Interaction.CombinedDCRatio}.Max()
    End Function
End Class

#End Region

#Region "Member Tension"
Partial Public Class tnxTowerOutputMemberTensionTowerSection
    Public Sub ConverttoEDSResults(tnx As tnxGeometry)
        Dim tnxSection As tnxGeometryRec = tnx.tnxSectionSelector(CInt(Me.Number))
        For Each compResultComponent In Me.ComponentType
            compResultComponent.AddMaxComponentResultstoSection(tnxSection)
        Next
    End Sub
End Class

Partial Public Class tnxTowerOutputMemberTensionComponentType
    Public Function MaxMember() As tnxTowerOutputMemberTensionMember
        Return Me.Member.MaxBy(Function(x) x.MaxRatio).FirstOrDefault
    End Function

    Public Sub AddMaxComponentResultstoSection(tnxSection As tnxGeometryRec)
        Dim controllingMember As tnxTowerOutputMemberTensionMember = Me.MaxMember()
        'Compression
        tnxSection.TNXResults.Add(New tnxResult(String.Format("ten_{0}_p", Me.Name.ToLower),
                                             controllingMember.Tension.PDCRatio,
                                             controllingMember.Tension.phiPn,
                                             controllingMember.Tension.Pu,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Bending X
        tnxSection.TNXResults.Add(New tnxResult(String.Format("ten_{0}_mx", Me.Name.ToLower),
                                             controllingMember.Bending.MxDCRatio,
                                             controllingMember.Bending.phiMnx,
                                             controllingMember.Bending.Mux,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Bending Y
        tnxSection.TNXResults.Add(New tnxResult(String.Format("ten_{0}_my", Me.Name.ToLower),
                                             controllingMember.Bending.MyDCRatio,
                                             controllingMember.Bending.phiMny,
                                             controllingMember.Bending.Muy,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Shear
        tnxSection.TNXResults.Add(New tnxResult(String.Format("ten_{0}_v", Me.Name.ToLower),
                                             controllingMember.Shear.VDCRatio,
                                             controllingMember.Shear.phiVn,
                                             controllingMember.Shear.Vu,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Torsion
        tnxSection.TNXResults.Add(New tnxResult(String.Format("ten_{0}_t", Me.Name.ToLower),
                                             controllingMember.Shear.TDCRatio,
                                             controllingMember.Shear.phiTn,
                                             controllingMember.Shear.Tu,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
        'Interaction
        tnxSection.TNXResults.Add(New tnxResult(String.Format("ten_{0}_int", Me.Name.ToLower),
                                             controllingMember.Interaction.CombinedDCRatio,
                                             Nothing,
                                             Nothing,
                                             Me.ComponentDesignParameters.DCRatioLimit,
                                             tnxSection))
    End Sub
End Class

Partial Public Class tnxTowerOutputMemberTensionMember
    Public Function MaxRatio() As Double
        Return {Tension.PDCRatio, Bending.MxDCRatio, Bending.MyDCRatio, Shear.VDCRatio, Shear.TDCRatio, Interaction.CombinedDCRatio}.Max()
    End Function
End Class
#End Region

#Region "Guys"
Partial Public Class tnxTowerOutputGuysTowerSection
    Public Sub ConverttoEDSResults(tnx As tnxGeometry)
        For Each guyMember In Member
            guyMember.AddGuyResultstoGuyRec(guyMember.tnxGuySelector(tnx))
        Next
    End Sub

End Class

Partial Public Class tnxTowerOutputGuyMember

    Public ReadOnly Property GuyLeg() As String
        Get
            'Guy result stores elevation and leg in one string called "LocationID" Example: <LocationID>88.00 (C) (622)</LocationID>
            Return Me.LocationID.Split("("c, ")"c)(1)
        End Get
    End Property

    Public ReadOnly Property GuyHeight() As Double
        Get
            'Guy result stores elevation and leg in one string called "LocationID" Example: <LocationID>88.00 (C) (622)</LocationID>
            Return CDbl(Me.LocationID.Split(" "c)(0))
        End Get
    End Property

    Public Sub AddGuyResultstoGuyRec(guyRec As tnxGuyRecord)
        If guyRec Is Nothing Then Exit Sub
        'Tension
        guyRec.TNXResults.Add(New tnxResult(String.Format("guy_{0}_tu", Me.GuyLeg.ToLower),
                                             1 / Me.ActualSF,
                                             Me.phiTn,
                                             Me.Tu,
                                             1 / Me.RequiredSF,
                                             guyRec))

    End Sub
    Public Function tnxGuySelector(tnx As tnxGeometry) As tnxGuyRecord
        'Select the guy record where results should be stored
        'You could have multiple guy recs with the same elevation and guy size so this is not gauranteed to work but it should atleast put the loads on a guy level at the same elevation
        Dim selectedGuy As List(Of tnxGuyRecord) = tnx.guyWires.Where(Function(x)
                                                                          Dim GuySize As String = ""
                                                                          Select Case Me.GuyLeg
                                                                              Case "A"
                                                                                  GuySize = x.GuySize
                                                                              Case "B"
                                                                                  GuySize = x.Guy120Size
                                                                              Case "C"
                                                                                  GuySize = x.Guy240Size
                                                                              Case "D"
                                                                                  GuySize = x.Guy360Size
                                                                          End Select

                                                                          Return x.GuyHeight = Me.GuyHeight AndAlso
                                                                    String.Format("{0} {1}", GuySize, x.GuyGrade) = Me.SizeDesignation
                                                                      End Function).ToList
        If selectedGuy.Count <> 1 Then
            Debug.Print(String.Format("Guy result not matched to guy record at Height: {0} Leg: {1}", Me.GuyHeight, Me.GuyLeg))
        End If
        Return selectedGuy.FirstOrDefault
    End Function
End Class
#End Region

#Region "Bolt Design"
Partial Public Class tnxTowerOutputBoltDesignTowerSection
    Public Sub ConverttoEDSResults(tnx As tnxGeometry, memberTension() As tnxTowerOutputMemberTensionTowerSection)
        'The list of memberTension tower sections is included to find the correct Ratio Limit (1.05%) for the bolts as it's not included in the current output
        Dim tnxSection As tnxGeometryRec = tnx.tnxSectionSelector(CInt(Me.Number))
        For Each compResultComponent In Me.ComponentType
            compResultComponent.AddMaxComponentResultstoSection(tnxSection, memberTension.Where(Function(x) x.Number = Me.Number).FirstOrDefault)
        Next
    End Sub
End Class

Partial Public Class tnxTowerOutputBoltDesignComponentType
    Public Function MaxMember() As tnxTowerOutputBoltDesignMember
        Return Me.Member.MaxBy(Function(x) x.BoltRatio).FirstOrDefault
    End Function

    Public Sub AddMaxComponentResultstoSection(tnxSection As tnxGeometryRec, memberTension As tnxTowerOutputMemberTensionTowerSection)
        'The memberTension is included to find the correct Ratio Limit (1.05%) for the bolts as it's not included in the current output
        Dim controllingBolt As tnxTowerOutputBoltDesignMember = Me.MaxMember()
        Dim tensionRatioLimit As Double = memberTension.ComponentType.Where(Function(x) x.Name = Me.Name).FirstOrDefault.ComponentDesignParameters.DCRatioLimit
        'Bolt
        tnxSection.TNXResults.Add(New tnxResult(String.Format("bolt_{0}_max", Me.Name.ToLower),
                                             controllingBolt.BoltRatio,
                                             controllingBolt.BoltCapacity,
                                             controllingBolt.BoltLoad,
                                             tensionRatioLimit,
                                             tnxSection))
    End Sub
End Class

Partial Public Class tnxTowerOutputBoltDesignMember

End Class
#End Region