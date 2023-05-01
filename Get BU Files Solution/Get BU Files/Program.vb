Imports System
Imports System.IO
Imports Microsoft.SharePoint.Client

Module Program
    Public toolList As New List(Of String) From {"CCIplate", "Drilled Pier Foundation", "Leg Reinforcement", "Pier and Pad", "SST Unit Base", "CCISeismic", "CCIpole", "Pile", ".eri", "Guyed Anchor"}
    Public fullLIst As New List(Of String) From
        {
            "(BS and EHS) Guy Temp Table NEW 6-17-2020",
            "60° Bent Plate with Plates",
            "60° Bent Plate with SR Inside Leg",
            "60deg Bent Plate with SR on Heel",
            "800874_Pier and Pad_to_Block Rev. H",
            "871294 - Inner Anchor Assembly Check",
            "90° Angle on SR or Pipe Analysis",
            "A325 Bolt Sizing Tool",
            "Additional Anchor Rods SST [Not Welded to BP]_Rev H -",
            "Additional Anchor Rods SST [SR Leg, Not Welded to BP]_Rev G -",
            "Additional Anchor Rods SST [SR Leg, Welded to BP]_Rev G -",
            "Additional Anchor Rods SST [Welded to BP]_Rev H -",
            "Additional Anchor Rods SST on same Bolt Circle_Rev H",
            "Additional Anchor Rods SST on same Bolt Circle_Universal Rev F & G",
            "Additional Anchor Rods with Bracket",
            "Additional Anchor Rods with Brackets_Rev H",
            "Additional Anchor Rods",
            "Additional SR Diagonals with Clips",
            "Anchor Rod Breakout Rev H",
            "Anchor Rod Interaction Check - TIA-222-G - Section 4.9.9 [Monopoles]",
            "Anchor Rod Pullout",
            "AR embedment - Rev H",
            "Bolt on Anchor Rod Bracket",
            "Bolt on AR Bracket",
            "Bolted Bridge Stiffener Reinforcement",
            "Bolted Connection",
            "Bolt-on Flange Bracket",
            "Bolt-On Guy Lug Assembly",
            "Bolt-On Shear Stop Assembly - CCI-PO-0910",
            "BV Splice Plate Check",
            "CaAa - Bridge Stiffeners - Version 1.0",
            "CCIconcretepole",
            "CCIFlagpole Tool",
            "CCIplate",
            "CCIpole",
            "CCISeismic",
            "CCIwoodpole",
            "Compression Member Strength Reduction Tool",
            "CON-FRM-10375 Modification Scope Form",
            "Diagonal Reinforcement Crosby Clips Rev.G",
            "Diagonal Reinforcement Crosby Clips Rev.H",
            "Doubler Plate",
            "Drilled Pier Foundation",
            "DSMC",
            "Eccentric HSS Steel Grade Beam Analysis",
            "Eccentrically Loaded HSS Steel Grade Beam Analysis",
            "Eccentrically Loaded W-Shape Steel Grade Beam Design",
            "Eccentricity at Flange - Bolt Check",
            "Even_Guy Anchor Shaft Combined Forces Check",
            "Flange Plate Bypass SST [Not Welded to Flange]",
            "Footpad Modification Design",
            "GAEP Anchor Assembly Design",
            "GAEP Pier Design",
            "GAEP Wing Plate Calcs",
            "Guy Temp Table NEW 6-17-2020",
            "Guyed Anchor Block Foundation",
            "HSS Steel Grade Beam Analysis",
            "Leg Reinforcement Tool",
            "Load Case Form",
            "Mapped Monopole Geometry",
            "Monopole properties calculator",
            "New Drilled Pier Anchor Rod Embedment v1.0.0",
            "Odd_Guy Anchor Shaft Combined Forces Chec",
            "Pier and Pad Foundation",
            "Pier Design",
            "Pile Foundation",
            "PiRod SR Leg Tension Check Rev H",
            "Plate Section Check",
            "Reaction Comparison Test",
            "Reason for Insufficiency",
            "Redundant Bolted Connection",
            "ReinforcedLegChannelCheck",
            "Revision Log",
            "Sleeve Splice Connection Check",
            "SP Additional Stiffness",
            "Splice Check",
            "Splice Plate Weld Check",
            "SST Anchor Rod Eccentric Load",
            "SST Unit Base Foundation",
            "Steel Grade Beam Analysis",
            "Thumbs",
            "Topo Calculator",
            "Transition Stiffeners",
            "Truss Leg",
            "U-washer [9.3.14]",
            "Vertical Rebar Embedment",
            "Weld Analysis-IC Method v1.1",
            "Welded Bridge Stiffener with Pipe Legs",
            "Welded Bridge Stiffener",
            "Welded Plate Bridge Stiffener"
        }

    Sub Main(args As String())
        Console.WriteLine("Process Start")

        Dim allFiles As New List(Of String)
        Dim i As Integer = 0
        allFiles.Add(HeaderRow)

        Using csvWriter As New StreamWriter("C:\Users\Imiller\Work Area\SAPI Testing\BU_WO Files_Updated.csv")
            csvWriter.WriteLine(HeaderRow)

            Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("C:\Users\Imiller\Crown Castle USA Inc\MT - ECS Tools Collaboration - SA Process Improvement - SA Process Improvement\Test Plan\Test Plan Creation\BU_WOs.csv")
                csvReader.TextFieldType = FileIO.FieldType.Delimited
                csvReader.SetDelimiters(",")

                Dim currentRow As String()
                While Not csvReader.EndOfData
                    Try
                        currentRow = csvReader.ReadFields()

                        For Each str As String In LoopThroughFiles(currentRow)
                            csvWriter.WriteLine(str)
                        Next
                        'allFiles.AddRange(LoopThroughFiles(currentRow))
                    Catch ex As Exception
                        Console.WriteLine("Line " & currentRow(1) & "_" & currentRow(0) & " is not valid and will be skipped.")
                    End Try
                    i += 1
                    'Console.WriteLine(i)
                End While
            End Using

        End Using
        'File.WriteAllLines("C:\Users\Imiller\Work Area\SAPI Testing\BU_WO Files_Updated.csv", allFiles.ToArray)

        Console.WriteLine("Process End")
    End Sub


    Public Function LoopThroughFiles(ByVal sitedata As String()) As List(Of String)

        Dim files As New List(Of String)

        Dim filepath As DirectoryInfo = New DirectoryInfo(sitedata(5))
        Dim oFolder As DirectoryInfo

        Dim topFolder As String = ""
        Dim folname As String = sitedata(5)
        Dim fname As String = ""
        Dim newrow As String = ""

        'Determine if the new BU WO folder format is used
        If filepath.Exists Then
            For Each subDir In filepath.GetDirectories
                If LCase(subDir.Name).Contains("prod") Then
                    topFolder = folname & "\" & subDir.Name

                    Exit For
                End If
            Next
        Else
            topFolder = folname & " - SA\Report"
        End If

        topFolder = "https://crowncastle.sharepoint.com/sites/MT-ECS-Vendors/Shared Documents/Forms/AllItems.aspx?id=%2Fsites%2FMT%2DECS%2DVendors%2FShared Documents%2FGeneral%2FB%26T Engineering %2D Structural Files&viewid=b782b2ce%2Dcdd5%2D4b09%2D867a%2D9b817268d100"

        For Each oFile As FileInfo In New DirectoryInfo(topFolder).GetFiles("*", SearchOption.AllDirectories)
            newrow = ""

            'Make sure the filename isn't blank and only includes excel files and eri files. (CAD, Mathcad, Word, PDF, misc files all filtered out
            If Not (oFile.Name = String.Empty) And (InStr(LCase(oFile.Name), ".xls") > 0 Or InStr(LCase(oFile.Name), "eri") > 0) And Not InStr(LCase(oFile.Name), ".eri.") > 0 Then

                Dim towerType As String = sitedata(6)
                Dim bu As String = sitedata(1)
                Dim wo As String = sitedata(2)
                Dim str As String = sitedata(0)

                newrow += """" & bu & ""","
                newrow += """" & str & ""","
                newrow += """" & wo & ""","
                newrow += """" & oFile.Name & ""","
                newrow += """" & oFile.FullName & ""","
                fname = LCase(oFile.Name)

                For i = 0 To fullLIst.Count - 1
                    newrow += """" & IIf(LCase(oFile.Name).Contains(LCase(fullLIst(i))), "X", "") & ""","
                Next
            End If

            If newrow <> String.Empty Then
                files.Add(newrow)
            End If
        Next

        Return files
    End Function

    Public Function HeaderRow() As String
        Dim newrow As String

        newrow += """bu"","
        newrow += """str"","
        newrow += """wo"","
        newrow += """filename"","
        newrow += """filepath"","


        For i = 0 To fullLIst.Count - 1
            newrow += """" & fullLIst(i).Replace(",", "") & ""","
        Next

        Return newrow

    End Function
End Module

