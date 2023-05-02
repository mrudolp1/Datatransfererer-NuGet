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
            Next

            Me.Text = Me.Text & " - " & Now.ToString
            ButtonclickToggle(Me.Cursor)
        End Sub

        Private Sub Export(sender As Object, e As EventArgs) Handles ToolStripSplitButton1.Click
            ButtonclickToggle(Me.Cursor)
            Dim dir As String = IIf(CType(frmMain.chkWorkLocal.Checked, Boolean) = True, lFolder, rFolder)
            Dim testid As Integer = CType(frmMain.testID.Text, Integer)
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
            ButtonclickToggle(Me.Cursor)
        End Sub

    End Class

End Namespace
