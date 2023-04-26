Public Class frmSummary
    Public Property myDs As DataSet

    Private Sub frmSummary_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

            newPg.Text = dt.TableName
            newPg.Controls.Add(dgv)
            tcData.TabPages.Add(newPg)
        Next

        Me.Text = Me.Text & " - " & Now.ToString
    End Sub
End Class