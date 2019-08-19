Public Class DataGridViewCustom
    Inherits System.Windows.Forms.DataGridView

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Me.GridColor = System.Drawing.SystemColors.Control
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        Me.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.BackgroundColor = System.Drawing.SystemColors.Control
        Me.ReadOnly = True
        Me.AllowUserToAddRows = False
        Me.AllowUserToDeleteRows = False
        Me.AllowUserToOrderColumns = True
        Me.AllowUserToResizeColumns = True
        Me.AllowUserToResizeRows = False
        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
