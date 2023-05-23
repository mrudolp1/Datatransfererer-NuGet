Imports System.Runtime.CompilerServices

Namespace UnitTesting
    Public Class frmSummary
        Public Property myDs As DataSet

        Private Sub frmSummary_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            ButtonclickToggle(Me.Cursor)
            Dim counter As Integer = 1

            For Each dt As DataTable In Me.myDs.Tables
                Dim newPg As New TabPage
                Dim dgv As New DataGridView

                With dgv
                    .DataSource = dt
                    .Dock = DockStyle.Fill
                    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    .RowHeadersVisible = False
                    .BackgroundColor = Color.White
                    .AllowUserToAddRows = False
                    .AllowUserToDeleteRows = False
                    .AllowUserToOrderColumns = False
                    .AllowUserToResizeColumns = False
                    .AllowUserToResizeRows = False
                End With

                newPg.Name = "P_" & counter
                dgv.Name = "DGV_" & counter

                newPg.Text = dt.TableName
                newPg.Controls.Add(dgv)
                tcData.TabPages.Add(newPg)

                counter += 1

                For Each dr As DataGridViewRow In dgv.Rows
                    For Each dc As DataGridViewColumn In dgv.Columns
                        If dc.Name.Contains("Status") Then
                            If dr.Cells(dc.Name).Value.ToString.ToLower = "fail" Then
                                dr.Cells(dc.Name).Style.BackColor = Color.Red
                            Else
                                dr.Cells(dc.Name).Style.BackColor = Color.Green
                            End If
                            Exit For
                        End If
                    Next
                Next
            Next

            Me.Text = Me.Text & " - " & Now.ToString
            ButtonclickToggle(Me.Cursor)
        End Sub

        Private Sub btn_Export(sender As Object, e As EventArgs) Handles ToolStripSplitButton1.Click
            ButtonclickToggle(Me.Cursor)
            Export()
            ButtonclickToggle(Me.Cursor)
        End Sub


        Public Sub Export()
            Dim dir As String = frmMain.dirUse 'IIf(CType(frmMain.chkWorkLocal.Checked, Boolean) = True, frmMain.lFolder, frmMain.rFolder)
            Dim testid As Integer = CType(frmMain.testCase, Integer)
            Dim testiteration As Integer = CType(frmMain.testIteration.Text, Integer)
            Dim filepath As String


            dir = dir & "\Test ID " & testid & "\Iteration " & testiteration & "\Results " & Me.Text.ToDirectoryString()
            If Not IO.Directory.Exists(dir) Then IO.Directory.CreateDirectory(dir)


            For Each dt As DataTable In myDs.Tables
                filepath = dir & "\" & dt.TableName.ToDirectoryString & ".csv"
                Try
                    dt.ToCSV(filepath)
                Catch
                End Try
            Next
        End Sub
    End Class

End Namespace
