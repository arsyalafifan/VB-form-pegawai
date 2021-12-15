<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class poppeg
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.BtnOk = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtNama = New System.Windows.Forms.TextBox()
        Me.Counter = New System.Windows.Forms.TextBox()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SuspendLayout()
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7})
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(187, 102)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(473, 152)
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'BtnOk
        '
        Me.BtnOk.Location = New System.Drawing.Point(585, 282)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(75, 23)
        Me.BtnOk.TabIndex = 1
        Me.BtnOk.Text = "Ok"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(184, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Nama"
        '
        'TxtNama
        '
        Me.TxtNama.Location = New System.Drawing.Point(251, 53)
        Me.TxtNama.Name = "TxtNama"
        Me.TxtNama.Size = New System.Drawing.Size(409, 20)
        Me.TxtNama.TabIndex = 3
        '
        'Counter
        '
        Me.Counter.Location = New System.Drawing.Point(187, 284)
        Me.Counter.Name = "Counter"
        Me.Counter.Size = New System.Drawing.Size(106, 20)
        Me.Counter.TabIndex = 4
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "NIP"
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Nama Pegawai"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Bagian"
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Tanggal Lahir"
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Pendidikan"
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Status"
        '
        'poppeg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Counter)
        Me.Controls.Add(Me.TxtNama)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BtnOk)
        Me.Controls.Add(Me.ListView1)
        Me.Name = "poppeg"
        Me.Text = "poppeg"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListView1 As ListView
    Friend WithEvents BtnOk As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents TxtNama As TextBox
    Friend WithEvents Counter As TextBox
    Friend WithEvents ColumnHeader1 As ColumnHeader
    Friend WithEvents ColumnHeader2 As ColumnHeader
    Friend WithEvents ColumnHeader3 As ColumnHeader
    Friend WithEvents ColumnHeader4 As ColumnHeader
    Friend WithEvents ColumnHeader5 As ColumnHeader
    Friend WithEvents ColumnHeader6 As ColumnHeader
    Friend WithEvents ColumnHeader7 As ColumnHeader
End Class
