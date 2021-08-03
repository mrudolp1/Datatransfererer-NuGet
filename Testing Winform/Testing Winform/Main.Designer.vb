Partial Public Class frmMain
    Inherits DevExpress.XtraEditors.XtraForm

    ''' <summary>
    ''' Required designer variable.
    ''' </summary>
    Private components As System.ComponentModel.IContainer = Nothing

    ''' <summary>
    ''' Clean up any resources being used.
    ''' </summary>
    ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso (components IsNot Nothing) Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

#Region "Windows Form Designer generated code"

    ''' <summary>
    ''' Required method for Designer support - do not modify
    ''' the contents of this method with the code editor.
    ''' </summary>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.sqltoexcel = New System.Windows.Forms.Button()
        Me.exceltosql = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'sqltoexcel
        '
        Me.sqltoexcel.Location = New System.Drawing.Point(36, 12)
        Me.sqltoexcel.Name = "sqltoexcel"
        Me.sqltoexcel.Size = New System.Drawing.Size(160, 52)
        Me.sqltoexcel.TabIndex = 0
        Me.sqltoexcel.Text = "Load from SQL / Save to Excel"
        Me.sqltoexcel.UseVisualStyleBackColor = True
        '
        'exceltosql
        '
        Me.exceltosql.Location = New System.Drawing.Point(36, 102)
        Me.exceltosql.Name = "exceltosql"
        Me.exceltosql.Size = New System.Drawing.Size(160, 52)
        Me.exceltosql.TabIndex = 1
        Me.exceltosql.Text = "Load from Excel / Save to SQL"
        Me.exceltosql.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(260, 14)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(214, 140)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'frmMain
        '
        Me.Appearance.BackColor = System.Drawing.Color.White
        Me.Appearance.Options.UseBackColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(496, 187)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.exceltosql)
        Me.Controls.Add(Me.sqltoexcel)
        Me.IconOptions.Image = CType(resources.GetObject("frmMain.IconOptions.Image"), System.Drawing.Image)
        Me.Name = "frmMain"
        Me.Text = "EDS & Excel Testing"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents sqltoexcel As Button
    Friend WithEvents exceltosql As Button
    Friend WithEvents PictureBox1 As PictureBox

#End Region

End Class
